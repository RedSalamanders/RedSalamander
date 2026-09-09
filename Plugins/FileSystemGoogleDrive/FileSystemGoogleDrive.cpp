#include "FileSystemGoogleDrive.h"

#include <algorithm>
#include <array>
#include <barrier>
#include <cctype>
#include <charconv>
#include <condition_variable>
#include <cstring>
#include <format>
#include <limits>
#include <map>
#include <memory>
#include <new>
#include <optional>
#include <thread>
#include <unordered_set>
#include <utility>

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include <curl/curl.h>

#include "ContentDigest.h"
#include "CurlProcessRuntime.h"
#include "DeleteOnCloseTemporaryFile.h"
#include "FileSystemGoogleDriveResources.h"
#include "Helpers.h"
#include "PaginationGuard.h"
#include "UriEncoding.h"
#include "YyjsonHelpers.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

namespace
{
constexpr wchar_t kPluginId[]      = L"builtin/file-system-gdrive";
constexpr wchar_t kPluginShortId[] = L"gdrive";
constexpr wchar_t kPluginAuthor[]  = L"RedSalamander";
constexpr wchar_t kPluginVersion[] = VERSINFO_PLUGIN_VERSION;

std::atomic<unsigned long> g_fileSystemGoogleDriveInstanceCount{0u};
std::atomic<bool> g_fileSystemGoogleDriveShutdownRequested{false};

[[nodiscard]] Common::CurlRuntime::ProcessLease& GetCurlRuntimeLease() noexcept
{
    static Common::CurlRuntime::ProcessLease lease;
    return lease;
}

[[nodiscard]] HRESULT InitializeSharedCurlRuntime() noexcept
{
    return curl_global_init(CURL_GLOBAL_DEFAULT) == CURLE_OK ? S_OK : E_FAIL;
}

void CleanupSharedCurlRuntime() noexcept
{
    curl_global_cleanup();
}

[[nodiscard]] const wchar_t* LocalizedPluginName() noexcept
{
    static const std::wstring name = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMGOOGLEDRIVE_NAME);
    return name.c_str();
}

[[nodiscard]] const wchar_t* LocalizedPluginDescription() noexcept
{
    static const std::wstring description = LoadStringResource(g_hInstance, IDS_FILESYSTEMGOOGLEDRIVE_DESCRIPTION);
    return description.c_str();
}

constexpr unsigned int kCommandIdOpenConnection = 1u;

constexpr char kTokenEndpoint[] = "https://oauth2.googleapis.com/token";

#if defined(ENABLE_TESTS)
std::mutex g_debugDriveOriginMutex;
std::string g_debugDriveOrigin; // `http://127.0.0.1:<port>` while a loopback fixture is active
std::atomic_bool g_debugDriveSyntheticConnection{false};

[[nodiscard]] std::string DebugDriveOrigin()
{
    std::lock_guard lock(g_debugDriveOriginMutex);
    return g_debugDriveOrigin;
}
#endif

[[nodiscard]] std::string ApiOriginUrl()
{
#if defined(ENABLE_TESTS)
    if (const std::string origin = DebugDriveOrigin(); ! origin.empty())
    {
        return origin;
    }
#endif
    return "https://www.googleapis.com";
}

[[nodiscard]] std::string TokenEndpointUrl()
{
#if defined(ENABLE_TESTS)
    if (const std::string origin = DebugDriveOrigin(); ! origin.empty())
    {
        return origin + "/token";
    }
#endif
    return kTokenEndpoint;
}

[[nodiscard]] std::string FilesEndpointUrl()
{
    return ApiOriginUrl() + "/drive/v3/files";
}

[[nodiscard]] std::string UploadFilesEndpointUrl()
{
    return ApiOriginUrl() + "/upload/drive/v3/files";
}

[[nodiscard]] std::string AboutEndpointUrl()
{
    return ApiOriginUrl() + "/drive/v3/about";
}

[[nodiscard]] std::string DrivesEndpointUrl()
{
    return ApiOriginUrl() + "/drive/v3/drives/";
}

constexpr char kFolderMimeType[]             = "application/vnd.google-apps.folder";
constexpr char kShortcutMimeType[]           = "application/vnd.google-apps.shortcut";
constexpr size_t kMaxJsonResponseBytes       = 16u * 1024u * 1024u;
constexpr unsigned int kMaxAuthorizedRetries = 3u;
constexpr uint64_t kMaxRetryDelayMs          = 5'000u;

constexpr char kSchemaJson[] = R"json(
{
  "version": 1,
  "title": "Google Drive",
  "fields": [
    {
      "key": "defaultClientId",
      "label": "Default OAuth client id",
      "type": "text",
      "default": "",
      "description": "Desktop OAuth client id used when a Connection Manager profile selects 'useDefaultClientId'. Leave empty only if every profile provides its own client id."
    },
    {
      "key": "connectTimeoutMs",
      "label": "Connect timeout (ms)",
      "type": "value",
      "default": 10000,
      "description": "TCP connect timeout used for Google HTTPS requests.",
      "min": 1,
      "max": 600000
    },
    {
      "key": "requestTimeoutMs",
      "label": "Stall timeout (ms)",
      "type": "value",
      "default": 30000,
      "description": "Abort requests that make no forward progress for this long.",
      "min": 1,
      "max": 600000
    },
    {
      "key": "pageSize",
      "label": "Page size",
      "type": "value",
      "default": 200,
      "description": "Max children fetched per Drive API request (1..1000).",
      "min": 1,
      "max": 1000
    }
  ]
}
)json";

constexpr char kCapabilitiesJson[] = R"json(
{
  "version": 2,
  "pathProfile": "google-drive",
  "rootId": "configured-drive-root",
  "operations": {
    "copy": true,
    "move": true,
    "nativeMove": true,
    "delete": true,
    "rename": true,
    "createDirectory": true,
    "properties": true,
    "read": true,
    "write": true,
    "recycle": true
  },
  "concurrency": {
    "copyMoveMax": 4,
    "deleteMax": 4,
    "deleteRecycleBinMax": 1
  },
  "transfer": {
    "export": { "copy": ["*"], "move": [] },
    "import": { "copy": ["*"], "move": [] }
  },
  "identity": { "object": "providerItemId", "revision": "version", "boundDelete": false, "conditionalDelete": false },
  "publication": { "exclusiveStage": false, "conditionalPublish": false, "committedSize": true },
  "links": { "preserveFileLink": false, "preserveDirectoryLink": false, "retargetInTree": false, "exactLinkRemoval": false },
  "metadata": { "motw": "reported-loss", "alternateStreams": "reported-loss", "extendedAttributes": "reported-loss", "sparse": "reported-loss", "efs": "reported-loss" },
  "verification": { "hostReadback": false, "providerProof": "writer-digest" },
  "cancellation": { "abort": false, "deadline": true, "routeClass": "providerWatchdog", "providerWatchdogTimeoutMs": configured-drive-watchdog-ms },
  "names": {
    "pathTextStableIdentity": true,
    "comparison": "ordinalCaseSensitive",
    "normalization": "none",
    "preferredSeparator": "/",
    "acceptedSeparators": ["/"],
    "casePreserving": true,
    "caseOnlyRename": "supported",
    "maxComponentUtf16": 255
  },
  "directories": { "model": "providerVirtual" }
}
)json";

struct CurlSlistDeleter
{
    void operator()(curl_slist* list) const noexcept
    {
        if (list)
        {
            curl_slist_free_all(list);
        }
    }
};

using unique_curl_slist = std::unique_ptr<curl_slist, CurlSlistDeleter>;

struct CurlEasyDeleter
{
    void operator()(CURL* handle) const noexcept
    {
        if (handle)
        {
            curl_easy_cleanup(handle);
        }
    }
};

using unique_curl_easy  = std::unique_ptr<CURL, CurlEasyDeleter>;
using unique_yyjson_doc = Common::Json::UniqueDocument;

struct HttpResponse
{
    long statusCode = 0;
    std::string body;
    std::string retryAfter;
    std::string location; // resumable upload session URL
    std::string range;    // acknowledged byte range of a resumable upload
};

// R0f-GDrive containment. The owning File Operations call keeps its options on the current thread
// (`DriveOperationOptionsScope` at every mutation entry point); every libcurl transfer issued under it
// polls the host's operation control from its progress callback (libcurl invokes it at least once per
// second even while a request waits) and returns the host's own verdict (ERROR_CANCELLED /
// ERROR_TIMEOUT). Every request is also bounded by the connect timeout and the hard request timeout.
thread_local const FileSystemOptions* t_driveOperationOptions = nullptr;

class DriveOperationOptionsScope final
{
public:
    explicit DriveOperationOptionsScope(const FileSystemOptions* options) noexcept : _previous(t_driveOperationOptions)
    {
        if (options != nullptr)
        {
            t_driveOperationOptions = options;
        }
    }
    ~DriveOperationOptionsScope()
    {
        t_driveOperationOptions = _previous;
    }
    DriveOperationOptionsScope(const DriveOperationOptionsScope&)            = delete;
    DriveOperationOptionsScope& operator=(const DriveOperationOptionsScope&) = delete;
    DriveOperationOptionsScope(DriveOperationOptionsScope&&)                 = delete;
    DriveOperationOptionsScope& operator=(DriveOperationOptionsScope&&)      = delete;

private:
    const FileSystemOptions* _previous = nullptr;
};

// A throttle backoff is a quiet point: Cancel or a passed deadline ends it early.
void SleepWithOperationControl(uint64_t delayMs) noexcept
{
    const ULONGLONG deadline = GetTickCount64() + delayMs;
    while (true)
    {
        const ULONGLONG now = GetTickCount64();
        if (now >= deadline || FAILED(FileSystemCheckOperationControl(t_driveOperationOptions)))
        {
            return;
        }
        Sleep(static_cast<DWORD>((std::min<ULONGLONG>)(deadline - now, 100ull)));
    }
}

#if defined(_DEBUG)
using DebugHttpRequestHook = HRESULT (*)(void* cookie,
                                         std::string_view method,
                                         std::string_view url,
                                         const std::vector<std::string>& headers,
                                         std::string_view body,
                                         HttpResponse& response) noexcept;

std::atomic<DebugHttpRequestHook> g_debugHttpRequestHook{nullptr};
std::atomic<void*> g_debugHttpRequestCookie{nullptr};
std::atomic_bool g_debugSuppressRetrySleep{false};

class DebugHttpRequestHookScope final
{
public:
    DebugHttpRequestHookScope(DebugHttpRequestHook hook, void* cookie) noexcept
    {
        g_debugHttpRequestCookie.store(cookie, std::memory_order_release);
        g_debugHttpRequestHook.store(hook, std::memory_order_release);
    }
    ~DebugHttpRequestHookScope()
    {
        g_debugHttpRequestHook.store(nullptr, std::memory_order_release);
        g_debugHttpRequestCookie.store(nullptr, std::memory_order_release);
    }

    DebugHttpRequestHookScope(const DebugHttpRequestHookScope&)            = delete;
    DebugHttpRequestHookScope& operator=(const DebugHttpRequestHookScope&) = delete;
};
#endif

[[nodiscard]] std::wstring Utf16FromUtf8(std::string_view text) noexcept
{
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
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

[[nodiscard]] std::optional<std::wstring> TryGetJsonString(yyjson_val* root, const char* key) noexcept
{
    const Common::Json::MemberResult<std::string_view> value = Common::Json::GetStringMember(root, key, Common::Json::MemberRequirement::Optional);
    return value.HasValue() ? std::optional<std::wstring>{Utf16FromUtf8(value.value)} : std::nullopt;
}

[[nodiscard]] std::optional<std::string> TryGetJsonUtf8String(yyjson_val* root, const char* key) noexcept
{
    const Common::Json::MemberResult<std::string_view> value = Common::Json::GetStringMember(root, key, Common::Json::MemberRequirement::Optional);
    return value.HasValue() ? std::optional<std::string>{value.value} : std::nullopt;
}

[[nodiscard]] std::optional<bool> TryGetJsonBool(yyjson_val* root, const char* key) noexcept
{
    const Common::Json::MemberResult<bool> value = Common::Json::GetBoolMember(root, key, Common::Json::MemberRequirement::Optional);
    return value.HasValue() ? std::optional<bool>{value.value} : std::nullopt;
}

[[nodiscard]] std::optional<uint64_t> TryGetJsonUInt(yyjson_val* root, const char* key) noexcept
{
    const Common::Json::MemberResult<uint64_t> value = Common::Json::GetUInt64Member(root,
                                                                                     key,
                                                                                     Common::Json::MemberRequirement::Optional,
                                                                                     Common::Json::NumericStringPolicy::Reject,
                                                                                     Common::Json::UnsignedIntegerPolicy::RequireUnsignedStorage);
    return value.HasValue() ? std::optional<uint64_t>{value.value} : std::nullopt;
}

[[nodiscard]] std::optional<uint64_t> TryGetJsonUInt64Flexible(yyjson_val* root, const char* key) noexcept
{
    const Common::Json::MemberResult<uint64_t> value =
        Common::Json::GetUInt64Member(root, key, Common::Json::MemberRequirement::Optional, Common::Json::NumericStringPolicy::Allow);
    return value.HasValue() ? std::optional<uint64_t>{value.value} : std::nullopt;
}

[[nodiscard]] std::string UrlEncodeUtf8(std::string_view text)
{
    return Common::Uri::PercentEncodeBytes(text);
}

void AppendQueryParam(std::string& url, std::string_view key, std::string_view value)
{
    url.push_back(url.find('?') == std::string::npos ? '?' : '&');
    url.append(key);
    url.push_back('=');
    url.append(UrlEncodeUtf8(value));
}

[[nodiscard]] HRESULT EnsureCurlInitialized() noexcept
{
    if (g_fileSystemGoogleDriveShutdownRequested.load(std::memory_order_acquire))
    {
        return HRESULT_FROM_WIN32(ERROR_SHUTDOWN_IN_PROGRESS);
    }

    return GetCurlRuntimeLease().Acquire(InitializeSharedCurlRuntime);
}

void TryCompleteGoogleDriveShutdown() noexcept
{
    if (! g_fileSystemGoogleDriveShutdownRequested.load(std::memory_order_acquire) ||
        g_fileSystemGoogleDriveInstanceCount.load(std::memory_order_acquire) != 0u)
    {
        return;
    }

    Common::CurlRuntime::ProcessLease& lease = GetCurlRuntimeLease();
    if (lease.IsAcquired())
    {
        static_cast<void>(lease.Release(CleanupSharedCurlRuntime));
    }
}

struct CurlResponseWriteState final
{
    HttpResponse* response = nullptr;
    size_t maxBytes        = 0u;
    bool tooLarge          = false;
};

size_t CurlWriteToString(char* data, size_t size, size_t nmemb, void* userData) noexcept
{
    if (size != 0u && nmemb > (std::numeric_limits<size_t>::max)() / size)
    {
        return 0u;
    }
    const size_t bytes = size * nmemb;
    if (bytes == 0u || ! userData)
    {
        return bytes;
    }

    auto* state = static_cast<CurlResponseWriteState*>(userData);
    if (! state->response || bytes > state->maxBytes - (std::min)(state->response->body.size(), state->maxBytes))
    {
        state->tooLarge = true;
        return 0u;
    }
    // Mandatory C callback boundary: curl must not see C++ exceptions.
    try
    {
        state->response->body.append(data, bytes);
        return bytes;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        return 0;
    }
}

size_t CurlCaptureResponseHeader(char* data, size_t size, size_t nmemb, void* userData) noexcept
{
    if (size != 0u && nmemb > (std::numeric_limits<size_t>::max)() / size)
    {
        return 0u;
    }
    const size_t bytes = size * nmemb;
    if (bytes == 0u || ! userData)
    {
        return bytes;
    }

    std::string_view line(data, bytes);
    const size_t colon = line.find(':');
    if (colon == std::string_view::npos)
    {
        return bytes;
    }
    const std::string_view name = line.substr(0, colon);
    const auto nameIs           = [&](std::string_view expected) noexcept
    {
        return name.size() == expected.size() && std::equal(expected.begin(), expected.end(), name.begin(), [](char left, char right) noexcept {
            return static_cast<char>(std::tolower(static_cast<unsigned char>(left))) == static_cast<char>(std::tolower(static_cast<unsigned char>(right)));
        });
    };
    std::string* target = nullptr;
    auto* response      = static_cast<HttpResponse*>(userData);
    if (nameIs("Retry-After"))
    {
        target = &response->retryAfter;
    }
    else if (nameIs("Location"))
    {
        target = &response->location;
    }
    else if (nameIs("Range"))
    {
        target = &response->range;
    }
    if (target == nullptr)
    {
        return bytes;
    }
    line.remove_prefix(colon + 1u);
    while (! line.empty() && (line.front() == ' ' || line.front() == '\t'))
    {
        line.remove_prefix(1u);
    }
    while (! line.empty() && (line.back() == '\r' || line.back() == '\n' || line.back() == ' ' || line.back() == '\t'))
    {
        line.remove_suffix(1u);
    }
    try
    {
        target->assign(line);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        return 0u;
    }
    return bytes;
}

struct CurlTransferDeadline final
{
    uint64_t deadlineTickMs          = 0u;
    HRESULT abortStatus              = S_OK;
    const FileSystemOptions* options = nullptr; // R0f-GDrive: the owning call's operation control
};

int CurlCheckTransferDeadline(void* userData, curl_off_t, curl_off_t, curl_off_t, curl_off_t) noexcept
{
    auto* deadline = static_cast<CurlTransferDeadline*>(userData);
    if (deadline == nullptr)
    {
        return 0;
    }
    if (deadline->options != nullptr)
    {
        if (const HRESULT control = FileSystemCheckOperationControl(deadline->options); FAILED(control))
        {
            deadline->abortStatus = control;
            return 1;
        }
    }
    if (deadline->deadlineTickMs != 0u && GetTickCount64() >= deadline->deadlineTickMs)
    {
        deadline->abortStatus = HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        return 1;
    }
    return 0;
}

[[nodiscard]] HRESULT MapCurlCodeToHresult(CURLcode code) noexcept
{
    if (code == CURLE_OK)
        return S_OK;
    else if (code == CURLE_OPERATION_TIMEDOUT)
        return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
    else if (code == CURLE_COULDNT_RESOLVE_HOST)
        return HRESULT_FROM_WIN32(ERROR_HOST_UNREACHABLE);
    else if (code == CURLE_COULDNT_CONNECT)
        return HRESULT_FROM_WIN32(ERROR_CONNECTION_REFUSED);
    else if (code == CURLE_REMOTE_ACCESS_DENIED)
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    else if (code == CURLE_PEER_FAILED_VERIFICATION || code == CURLE_SSL_CONNECT_ERROR)
        return HRESULT_FROM_WIN32(ERROR_TRUST_FAILURE);
    else
        return E_FAIL;
}

[[nodiscard]] HRESULT MapHttpStatusToHresult(long statusCode) noexcept
{
    if (statusCode >= 200 && statusCode < 300)
    {
        return S_OK;
    }

    switch (statusCode)
    {
        case 400: return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        case 401: return HRESULT_FROM_WIN32(ERROR_NOT_AUTHENTICATED);
        case 403: return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        case 404: return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        case 409: return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        case 429: return HRESULT_FROM_WIN32(ERROR_RETRY);
        default:
            if (statusCode >= 500 && statusCode <= 599)
            {
                return HRESULT_FROM_WIN32(ERROR_RETRY);
            }
            return E_FAIL;
    }
}

HRESULT PerformHttpRequest(std::string_view method,
                           std::string_view url,
                           const std::vector<std::string>& headers,
                           std::string_view body,
                           uint32_t connectTimeoutMs,
                           uint32_t requestTimeoutMs,
                           HttpResponse& response,
                           size_t maxResponseBytes = kMaxJsonResponseBytes) noexcept
{
    response = {};

#if defined(_DEBUG)
    if (const DebugHttpRequestHook hook = g_debugHttpRequestHook.load(std::memory_order_acquire))
    {
        const HRESULT hookHr = hook(g_debugHttpRequestCookie.load(std::memory_order_acquire), method, url, headers, body, response);
        if (SUCCEEDED(hookHr) && response.body.size() > maxResponseBytes)
        {
            response.body.clear();
            return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }
        return hookHr;
    }
#endif

    const HRESULT initHr = EnsureCurlInitialized();
    if (FAILED(initHr))
    {
        return initHr;
    }

    unique_curl_easy curl(curl_easy_init());
    if (! curl)
    {
        return E_OUTOFMEMORY;
    }

    unique_curl_slist curlHeaders;
    for (const std::string& header : headers)
    {
        curl_slist* next = curl_slist_append(curlHeaders.get(), header.c_str());
        if (! next)
        {
            return E_OUTOFMEMORY;
        }
        curlHeaders.release();
        curlHeaders.reset(next);
    }

    const std::string urlCopy(url);
    CurlResponseWriteState writeState{.response = &response, .maxBytes = maxResponseBytes};
    CurlTransferDeadline transferDeadline{
        .deadlineTickMs = requestTimeoutMs == 0u ? 0u : Common::Paging::DeadlineFromNow(GetTickCount64(), requestTimeoutMs),
    };
    transferDeadline.options = t_driveOperationOptions;
    curl_easy_setopt(curl.get(), CURLOPT_URL, urlCopy.c_str());
    curl_easy_setopt(curl.get(), CURLOPT_NOSIGNAL, 1L);
    curl_easy_setopt(curl.get(), CURLOPT_ACCEPT_ENCODING, "");
    curl_easy_setopt(curl.get(), CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_2TLS);
    curl_easy_setopt(curl.get(), CURLOPT_USERAGENT, "RedSalamander/GoogleDrive/0.1");
    curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, &CurlWriteToString);
    curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, &writeState);
    curl_easy_setopt(curl.get(), CURLOPT_HEADERFUNCTION, &CurlCaptureResponseHeader);
    curl_easy_setopt(curl.get(), CURLOPT_HEADERDATA, &response);
    curl_easy_setopt(curl.get(), CURLOPT_NOPROGRESS, 0L);
    curl_easy_setopt(curl.get(), CURLOPT_XFERINFOFUNCTION, &CurlCheckTransferDeadline);
    curl_easy_setopt(curl.get(), CURLOPT_XFERINFODATA, &transferDeadline);

    if (connectTimeoutMs > 0)
    {
        curl_easy_setopt(curl.get(), CURLOPT_CONNECTTIMEOUT_MS, static_cast<long>(connectTimeoutMs));
    }

    if (requestTimeoutMs > 0)
    {
        curl_easy_setopt(curl.get(), CURLOPT_TIMEOUT_MS, static_cast<long>(requestTimeoutMs));
        const long lowSpeedTime = static_cast<long>((requestTimeoutMs + 999u) / 1000u);
        curl_easy_setopt(curl.get(), CURLOPT_LOW_SPEED_LIMIT, 1L);
        curl_easy_setopt(curl.get(), CURLOPT_LOW_SPEED_TIME, (std::max)(1L, lowSpeedTime));
    }

    if (curlHeaders)
    {
        curl_easy_setopt(curl.get(), CURLOPT_HTTPHEADER, curlHeaders.get());
    }

    const std::string methodCopy(method);
    std::string bodyCopy;
    if (method == "POST" || method == "PUT" || method == "PATCH")
    {
        bodyCopy.assign(body);
        if (method == "POST")
        {
            curl_easy_setopt(curl.get(), CURLOPT_POST, 1L);
        }
        else
        {
            curl_easy_setopt(curl.get(), CURLOPT_CUSTOMREQUEST, methodCopy.c_str());
        }
        curl_easy_setopt(curl.get(), CURLOPT_POSTFIELDS, bodyCopy.c_str());
        curl_easy_setopt(curl.get(), CURLOPT_POSTFIELDSIZE_LARGE, static_cast<curl_off_t>(bodyCopy.size()));
    }
    else if (method == "DELETE")
    {
        if (! body.empty())
        {
            return E_INVALIDARG;
        }
        curl_easy_setopt(curl.get(), CURLOPT_CUSTOMREQUEST, methodCopy.c_str());
    }
    else
    {
        if (! body.empty())
        {
            return E_INVALIDARG;
        }

        curl_easy_setopt(curl.get(), CURLOPT_HTTPGET, 1L);
    }

    const CURLcode code = curl_easy_perform(curl.get());
    if (code != CURLE_OK)
    {
        if (writeState.tooLarge)
        {
            response.body.clear();
            return HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        }
        if (FAILED(transferDeadline.abortStatus))
        {
            return transferDeadline.abortStatus;
        }
        return MapCurlCodeToHresult(code);
    }

    long statusCode = 0;
    curl_easy_getinfo(curl.get(), CURLINFO_RESPONSE_CODE, &statusCode);
    response.statusCode = statusCode;
    return S_OK;
}

[[nodiscard]] bool TryParseDigits(std::string_view text, size_t offset, size_t count, unsigned int& value) noexcept
{
    value = 0;
    if (offset + count > text.size())
    {
        return false;
    }

    for (size_t index = 0; index < count; ++index)
    {
        const char ch = text[offset + index];
        if (ch < '0' || ch > '9')
        {
            return false;
        }

        value = (value * 10u) + static_cast<unsigned int>(ch - '0');
    }

    return true;
}

[[nodiscard]] __int64 ParseRfc3339ToFileTime64(std::string_view text) noexcept
{
    if (text.size() < 20u)
    {
        return 0;
    }

    unsigned int year   = 0;
    unsigned int month  = 0;
    unsigned int day    = 0;
    unsigned int hour   = 0;
    unsigned int minute = 0;
    unsigned int second = 0;

    if (! TryParseDigits(text, 0u, 4u, year) || text[4] != '-' || ! TryParseDigits(text, 5u, 2u, month) || text[7] != '-' ||
        ! TryParseDigits(text, 8u, 2u, day) || (text[10] != 'T' && text[10] != 't') || ! TryParseDigits(text, 11u, 2u, hour) || text[13] != ':' ||
        ! TryParseDigits(text, 14u, 2u, minute) || text[16] != ':' || ! TryParseDigits(text, 17u, 2u, second))
    {
        return 0;
    }

    size_t offset       = 19u;
    unsigned int millis = 0;
    if (offset < text.size() && text[offset] == '.')
    {
        ++offset;
        unsigned int scale = 100u;
        while (offset < text.size() && text[offset] >= '0' && text[offset] <= '9')
        {
            if (scale > 0u)
            {
                millis += static_cast<unsigned int>(text[offset] - '0') * scale;
                scale /= 10u;
            }
            ++offset;
        }
    }

    int offsetMinutes = 0;
    if (offset >= text.size())
    {
        return 0;
    }

    if (text[offset] == 'Z' || text[offset] == 'z')
    {
        ++offset;
    }
    else if (text[offset] == '+' || text[offset] == '-')
    {
        const bool positive = text[offset] == '+';
        ++offset;

        unsigned int offsetHours = 0;
        unsigned int offsetMins  = 0;
        if (! TryParseDigits(text, offset, 2u, offsetHours))
        {
            return 0;
        }
        offset += 2u;

        if (offset < text.size() && text[offset] == ':')
        {
            ++offset;
        }

        if (! TryParseDigits(text, offset, 2u, offsetMins))
        {
            return 0;
        }
        offset += 2u;

        offsetMinutes = static_cast<int>(offsetHours * 60u + offsetMins);
        if (positive)
        {
            offsetMinutes = -offsetMinutes;
        }
    }
    else
    {
        return 0;
    }

    if (offset != text.size())
    {
        return 0;
    }

    SYSTEMTIME st{};
    st.wYear         = static_cast<WORD>(year);
    st.wMonth        = static_cast<WORD>(month);
    st.wDay          = static_cast<WORD>(day);
    st.wHour         = static_cast<WORD>(hour);
    st.wMinute       = static_cast<WORD>(minute);
    st.wSecond       = static_cast<WORD>(second);
    st.wMilliseconds = static_cast<WORD>(millis);

    FILETIME ft{};
    if (SystemTimeToFileTime(&st, &ft) == 0)
    {
        return 0;
    }

    ULARGE_INTEGER value{};
    value.LowPart  = ft.dwLowDateTime;
    value.HighPart = ft.dwHighDateTime;

    if (offsetMinutes != 0)
    {
        const int64_t adjustment = static_cast<int64_t>(offsetMinutes) * 60ll * 10'000'000ll;
        const int64_t adjusted   = static_cast<int64_t>(value.QuadPart) + adjustment;
        if (adjusted < 0)
        {
            return 0;
        }
        value.QuadPart = static_cast<uint64_t>(adjusted);
    }

    if (value.QuadPart > static_cast<uint64_t>((std::numeric_limits<__int64>::max)()))
    {
        return static_cast<__int64>((std::numeric_limits<__int64>::max)());
    }

    return static_cast<__int64>(value.QuadPart);
}

[[nodiscard]] std::wstring MakeSyntheticDisplayName(std::wstring_view name, std::wstring_view id)
{
    std::wstring encodedId;
    if (! Common::Uri::TryPercentEncodeUtf8ToWide(id, Common::Uri::SlashPolicy::Encode, encodedId))
    {
        encodedId.assign(id);
    }
    if (name.empty())
    {
        return std::format(L"[id:{}]", encodedId);
    }

    return std::format(L"{} [id:{}]", name, encodedId);
}

[[nodiscard]] bool IsRetryableAuthorizedStatus(long statusCode) noexcept
{
    return statusCode == 429 || (statusCode >= 500 && statusCode <= 599);
}

[[nodiscard]] uint64_t ComputeRetryDelayMs(const HttpResponse& response, unsigned int retryIndex) noexcept
{
    uint64_t retryAfterSeconds = 0u;
    if (! response.retryAfter.empty())
    {
        const auto parsed = std::from_chars(response.retryAfter.data(), response.retryAfter.data() + response.retryAfter.size(), retryAfterSeconds);
        if (parsed.ec == std::errc{} && parsed.ptr == response.retryAfter.data() + response.retryAfter.size())
        {
            return (std::min)(retryAfterSeconds > kMaxRetryDelayMs / 1000u ? kMaxRetryDelayMs : retryAfterSeconds * 1000u, kMaxRetryDelayMs);
        }
    }

    const uint64_t exponential = (std::min)(250ull << (std::min)(retryIndex, 4u), kMaxRetryDelayMs);
    const uint64_t jitterSpan  = exponential / 4u;
    return (std::min)(exponential + (jitterSpan == 0u ? 0u : GetTickCount64() % (jitterSpan + 1u)), kMaxRetryDelayMs);
}

[[nodiscard]] bool LooksLikeSyntheticDisplayName(std::wstring_view name) noexcept
{
    constexpr std::wstring_view kMarker = L" [id:";
    return name.ends_with(L"]") && (name.starts_with(L"[id:") || name.rfind(kMarker) != std::wstring_view::npos);
}

[[nodiscard]] std::wstring MakeExposedItemName(std::wstring_view name, std::wstring_view id, size_t exactSiblingNameCount)
{
    return exactSiblingNameCount > 1u || name.empty() || LooksLikeSyntheticDisplayName(name) ? MakeSyntheticDisplayName(name, id) : std::wstring(name);
}

[[nodiscard]] std::wstring ConnectionKeyForCache(std::wstring_view connectionName, std::wstring_view clientId)
{
    return std::format(L"{}|{}", connectionName, clientId);
}

[[nodiscard]] std::wstring BuildDriveDisplayName(std::wstring_view connectionName, std::wstring_view path)
{
    if (path.empty() || path == L"/")
    {
        return std::format(L"gdrive://{}", connectionName);
    }

    return std::format(L"gdrive://{}{}", connectionName, path);
}

template <typename TCallable> HRESULT RunBoundary(const wchar_t* operation, TCallable&& callable) noexcept
{
    try
    {
        return callable();
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"GDrive: {} failed with std::exception", operation);
        return E_FAIL;
    }
}
} // namespace

struct FileSystemGoogleDrive::ResolvedConnection
{
    ResolvedConnection() = default;
    ~ResolvedConnection()
    {
        SecureWipe::SecureClear(refreshToken);
    }
    ResolvedConnection(const ResolvedConnection&)                = delete;
    ResolvedConnection& operator=(const ResolvedConnection&)     = delete;
    ResolvedConnection(ResolvedConnection&&) noexcept            = default;
    ResolvedConnection& operator=(ResolvedConnection&&) noexcept = default;

    std::wstring connectionName;
    std::wstring canonicalPath = L"/";
    std::wstring connectionKey;
    std::wstring clientId;
    std::wstring refreshToken;
    std::wstring rootKind = L"myDrive";
    std::wstring sharedDriveId;
    std::wstring googleDocsMode = L"native";
    bool readOnly               = false;
    uint32_t connectTimeoutMs   = 10'000;
    uint32_t requestTimeoutMs   = 30'000;
    unsigned long pageSize      = 200;
};

struct FileSystemGoogleDrive::GoogleItem
{
    std::wstring id;
    std::wstring name;
    std::wstring mimeType;
    std::wstring version; // Drive's monotonically increasing revision counter for the file
    unsigned long attributes = FILE_ATTRIBUTE_NORMAL;
    uint64_t sizeBytes       = 0;
    __int64 creationTime     = 0;
    __int64 lastAccessTime   = 0;
    __int64 lastWriteTime    = 0;
    __int64 changeTime       = 0;
    bool isFolder            = false;
    bool isShortcut          = false;
    std::string sha256Hex; // R3-2: Drive's own SHA-256 of the binary content (empty for native documents)
};

struct FileSystemGoogleDrive::DriveInfoPayload
{
    std::wstring userDisplayName;
    std::wstring userEmail;
    std::wstring driveName;
    uint64_t totalBytes = 0;
    uint64_t freeBytes  = 0;
    uint64_t usedBytes  = 0;
    bool hasTotal       = false;
    bool hasFree        = false;
    bool hasUsed        = false;
};

FileSystemGoogleDrive::FileSystemGoogleDrive(IHost* host)
{
    g_fileSystemGoogleDriveInstanceCount.fetch_add(1u, std::memory_order_relaxed);
    _metaData.id          = kPluginId;
    _metaData.shortId     = kPluginShortId;
    _metaData.name        = LocalizedPluginName();
    _metaData.description = LocalizedPluginDescription();
    _metaData.author      = kPluginAuthor;
    _metaData.version     = kPluginVersion;

    if (host)
    {
        static_cast<void>(host->QueryInterface(__uuidof(IHostAlerts), _hostAlerts.put_void()));
        static_cast<void>(host->QueryInterface(__uuidof(IHostConnections), _hostConnections.put_void()));
    }
}

FileSystemGoogleDrive::~FileSystemGoogleDrive()
{
    {
        std::scoped_lock lock(_tokenMutex);
        _accessTokensByConnectionKey.clear();
    }
    g_fileSystemGoogleDriveInstanceCount.fetch_sub(1u, std::memory_order_acq_rel);
    TryCompleteGoogleDriveShutdown();
}

namespace FileSystemGoogleDriveInternal
{
bool CanCreateInstance() noexcept
{
    return ! g_fileSystemGoogleDriveShutdownRequested.load(std::memory_order_acquire);
}

void BeginShutdown() noexcept
{
    g_fileSystemGoogleDriveShutdownRequested.store(true, std::memory_order_release);
    TryCompleteGoogleDriveShutdown();
}

bool CanUnloadNow() noexcept
{
    TryCompleteGoogleDriveShutdown();
    return g_fileSystemGoogleDriveShutdownRequested.load(std::memory_order_acquire) &&
           g_fileSystemGoogleDriveInstanceCount.load(std::memory_order_acquire) == 0u && ! GetCurlRuntimeLease().IsAcquired();
}

#if defined(_DEBUG)
HRESULT RunDebugCurlRuntimeProbe() noexcept
{
    const HRESULT initializeHr = EnsureCurlInitialized();
    if (FAILED(initializeHr))
    {
        return initializeHr;
    }

    unique_curl_easy handle(curl_easy_init());
    return handle ? S_OK : E_OUTOFMEMORY;
}
#endif
} // namespace FileSystemGoogleDriveInternal

IHostAlerts* FileSystemGoogleDrive::GetHostAlerts() const noexcept
{
    return _hostAlerts.get();
}

void FileSystemGoogleDrive::ShowMissingClientIdAlert() const noexcept
{
    IHostAlerts* const hostAlerts = GetHostAlerts();
    if (! hostAlerts)
    {
        return;
    }

    std::wstring pluginName = _metaData.name ? _metaData.name : LocalizedPluginName();
    if (pluginName.empty())
    {
        pluginName = L"Google Drive";
    }

    const std::wstring title   = LoadStringResource(g_hInstance, IDS_FILESYSTEMGOOGLEDRIVE_ALERT_TITLE_SIGNIN_CONFIG_REQUIRED);
    const std::wstring message = FormatStringResource(g_hInstance, IDS_FILESYSTEMGOOGLEDRIVE_ALERT_MSG_MISSING_CLIENT_ID_FMT, pluginName);
    if (message.empty())
    {
        return;
    }

    HostAlertRequest request{};
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_APPLICATION;
    request.modality     = HOST_ALERT_MODAL;
    request.severity     = HOST_ALERT_ERROR;
    request.targetWindow = nullptr;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(hostAlerts->ShowAlert(&request, nullptr));
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (! ppvObject)
    {
        return E_POINTER;
    }

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
    {
        *ppvObject = static_cast<IFileSystem*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemPathCapabilities2))
    {
        *ppvObject = static_cast<IFileSystemPathCapabilities2*>(static_cast<IFileSystem*>(this));
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemRouteCapabilities))
    {
        *ppvObject = static_cast<IFileSystemRouteCapabilities*>(static_cast<FileSystemRouteCapabilitiesBase*>(this));
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemIO))
    {
        *ppvObject = static_cast<IFileSystemIO*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemDirectoryOperations))
    {
        *ppvObject = static_cast<IFileSystemDirectoryOperations*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemAtomicWriter))
    {
        *ppvObject = static_cast<IFileSystemAtomicWriter*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemIdentityDelete))
    {
        *ppvObject = static_cast<IFileSystemIdentityDelete*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(INavigationMenu))
    {
        *ppvObject = static_cast<INavigationMenu*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IDriveInfo))
    {
        *ppvObject = static_cast<IDriveInfo*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE FileSystemGoogleDrive::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE FileSystemGoogleDrive::Release() noexcept
{
    const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (result == 0)
    {
        delete this;
    }
    return result;
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (! metaData)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (! schemaJsonUtf8)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = StaticConfigurationSchema();
    return S_OK;
}

const char* GetFileSystemGoogleDriveStaticConfigurationSchema() noexcept
{
    return FileSystemGoogleDrive::StaticConfigurationSchema();
}

const char* FileSystemGoogleDrive::StaticConfigurationSchema() noexcept
{
    return kSchemaJson;
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    return RunBoundary(L"SetConfiguration", [&]() { return SetConfigurationImpl(configurationJsonUtf8); });
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (! configurationJsonUtf8)
    {
        return E_POINTER;
    }

    std::scoped_lock lock(_stateMutex);
    *configurationJsonUtf8 = _configurationJsonStorage[_configurationJsonIndex].c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (! pSomethingToSave)
    {
        return E_POINTER;
    }

    std::scoped_lock lock(_stateMutex);
    const auto& config = _configurationJsonStorage[_configurationJsonIndex];
    *pSomethingToSave  = (! config.empty() && config != "{}") ? TRUE : FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetMenuItems(const NavigationMenuItem** items, unsigned int* count) noexcept
{
    return RunBoundary(L"GetMenuItems",
                       [&]() -> HRESULT
    {
        if (! items || ! count)
        {
            return E_POINTER;
        }

        std::scoped_lock lock(_stateMutex);

        _menuEntries.clear();
        _menuEntryView.clear();

        MenuEntry header;
        header.flags = NAV_MENU_ITEM_FLAG_HEADER;
        header.label = _metaData.name ? _metaData.name : L"";
        _menuEntries.push_back(std::move(header));

        MenuEntry separator;
        separator.flags = NAV_MENU_ITEM_FLAG_SEPARATOR;
        _menuEntries.push_back(std::move(separator));

        MenuEntry openConnection;
        openConnection.label     = LoadStringResource(nullptr, IDS_MENU_CONNECTIONS_ELLIPSIS);
        openConnection.commandId = kCommandIdOpenConnection;
        _menuEntries.push_back(std::move(openConnection));

        MenuEntry root;
        root.label = L"/";
        root.path  = L"/";
        _menuEntries.push_back(std::move(root));

        _menuEntryView.reserve(_menuEntries.size());
        for (const auto& entry : _menuEntries)
        {
            NavigationMenuItem item{};
            item.flags     = entry.flags;
            item.label     = entry.label.empty() ? nullptr : entry.label.c_str();
            item.path      = entry.path.empty() ? nullptr : entry.path.c_str();
            item.iconPath  = entry.iconPath.empty() ? nullptr : entry.iconPath.c_str();
            item.commandId = entry.commandId;
            _menuEntryView.push_back(item);
        }

        *items = _menuEntryView.empty() ? nullptr : _menuEntryView.data();
        *count = static_cast<unsigned int>(_menuEntryView.size());
        return S_OK;
    });
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::ExecuteMenuCommand(unsigned int commandId) noexcept
{
    return RunBoundary(L"ExecuteMenuCommand",
                       [&]() -> HRESULT
    {
        if (commandId != kCommandIdOpenConnection)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
        }

        wil::com_ptr<IHostConnections> hostConnections = _hostConnections;
        if (! hostConnections)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        NavigationMenuCallbackSnapshot callbackSnapshot{};
        if (! TryCaptureNavigationMenuCallback(callbackSnapshot))
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        HostConnectionManagerRequest request{};
        request.sizeBytes      = sizeof(request);
        request.filterPluginId = kPluginId;
        request.ownerWindow    = nullptr;

        HostConnectionManagerResult result{};
        result.sizeBytes = sizeof(result);

        const HRESULT showHr = hostConnections->ShowConnectionManager(&request, &result);
        if (FAILED(showHr) || showHr == S_FALSE)
        {
            return showHr;
        }

        wil::unique_cotaskmem_string connectionName(result.connectionName);
        if (! connectionName || connectionName.get()[0] == L'\0')
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const std::wstring targetPath = std::format(L"/@conn:{}/", connectionName.get());
        return InvokeNavigationMenuCallback(callbackSnapshot, targetPath.c_str());
    });
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::SetCallback(INavigationMenuCallback* callback, void* cookie) noexcept
{
    _navigationMenuCallbackState.Set(callback, cookie);
    return S_OK;
}

bool FileSystemGoogleDrive::TryCaptureNavigationMenuCallback(NavigationMenuCallbackSnapshot& snapshot) noexcept
{
    return _navigationMenuCallbackState.TryCapture(snapshot);
}

HRESULT FileSystemGoogleDrive::InvokeNavigationMenuCallback(const NavigationMenuCallbackSnapshot& snapshot, const wchar_t* path) noexcept
{
    INavigationMenuCallback* callback = nullptr;
    void* cookie                      = nullptr;
    if (! _navigationMenuCallbackState.TryEnter(snapshot, callback, cookie))
    {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    }

    auto finishInvoke = wil::scope_exit([this]() noexcept { _navigationMenuCallbackState.FinishInvoke(); });
    return callback->NavigationMenuRequestNavigate(path, cookie);
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetDriveInfo(const wchar_t* path, DriveInfo* info) noexcept
{
    return RunBoundary(L"GetDriveInfo", [&]() { return GetDriveInfoImpl(path, info); });
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetDriveMenuItems(const wchar_t* /*path*/, const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (! items || ! count)
    {
        return E_POINTER;
    }

    *items = nullptr;
    *count = 0;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::ExecuteDriveMenuCommand(unsigned int /*commandId*/, const wchar_t* /*path*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept
{
    return RunBoundary(L"ReadDirectoryInfo", [&]() { return ReadDirectoryInfoImpl(path, ppFilesInformation); });
}

// ---------------------------------------------------------------------------------------------
// R0f-GDrive: Google Drive as a File Operations destination. Every entry point pins the task's
// options on the thread (`DriveOperationOptionsScope`) so each libcurl transfer polls the host's
// operation control, resolves the exposed path text to exactly one Drive object (duplicate names
// carry the `[id:...]` decoration and an ambiguous text fails closed with ERROR_DUP_NAME), performs
// the mutation through the Drive v3 API, and reports a mutation receipt for every item completion.
// ---------------------------------------------------------------------------------------------
namespace
{
constexpr std::string_view kDriveItemFields = "id,name,mimeType,modifiedTime,size,trashed,version,sha256Checksum";
constexpr uint64_t kDriveUploadChunkBytes   = 8ull * 1024u * 1024u; // a multiple of 256 KiB, as Drive requires

[[nodiscard]] bool IsDriveNotFound(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
}

[[nodiscard]] std::string JsonQuote(std::string_view text)
{
    std::string out;
    out.reserve(text.size() + 2u);
    out.push_back('"');
    for (const char ch : text)
    {
        switch (ch)
        {
            case '"': out += "\\\""; break;
            case '\\': out += "\\\\"; break;
            case '\n': out += "\\n"; break;
            case '\r': out += "\\r"; break;
            case '\t': out += "\\t"; break;
            default:
                if (static_cast<unsigned char>(ch) < 0x20u)
                {
                    out += std::format("\\u{:04x}", static_cast<unsigned>(static_cast<unsigned char>(ch)));
                }
                else
                {
                    out.push_back(ch);
                }
                break;
        }
    }
    out.push_back('"');
    return out;
}

[[nodiscard]] std::string JsonQuote(std::wstring_view text)
{
    return JsonQuote(std::string_view(Utf8FromUtf16(text)));
}

// Parses one Drive `files` resource. A trashed object is reported as not found: the plugin never
// exposes the trash, so a trashed object must not be addressable through its former path.
[[nodiscard]] HRESULT ParseGoogleItem(yyjson_val* entry, FileSystemGoogleDrive::GoogleItem& item)
{
    item = {};
    if (! entry || ! yyjson_is_obj(entry))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    if (TryGetJsonBool(entry, "trashed").value_or(false))
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    const auto idUtf8   = TryGetJsonUtf8String(entry, "id");
    const auto nameWide = TryGetJsonString(entry, "name");
    const auto mimeUtf8 = TryGetJsonUtf8String(entry, "mimeType");
    if (! idUtf8.has_value() || idUtf8->empty() || ! mimeUtf8.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    item.id         = Utf16FromUtf8(idUtf8.value());
    item.name       = nameWide.value_or(Utf16FromUtf8(idUtf8.value()));
    item.mimeType   = Utf16FromUtf8(mimeUtf8.value());
    item.isFolder   = mimeUtf8.value() == kFolderMimeType;
    item.isShortcut = mimeUtf8.value() == kShortcutMimeType;
    item.attributes = item.isFolder ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;
    if (item.isShortcut)
    {
        item.attributes |= FILE_ATTRIBUTE_REPARSE_POINT;
    }
    if (const auto sizeBytes = TryGetJsonUInt64Flexible(entry, "size"); sizeBytes.has_value())
    {
        item.sizeBytes = sizeBytes.value();
    }
    if (const auto modifiedTime = TryGetJsonUtf8String(entry, "modifiedTime"); modifiedTime.has_value())
    {
        item.lastWriteTime  = ParseRfc3339ToFileTime64(modifiedTime.value());
        item.creationTime   = item.lastWriteTime;
        item.lastAccessTime = item.lastWriteTime;
        item.changeTime     = item.lastWriteTime;
    }
    if (const auto version = TryGetJsonString(entry, "version"); version.has_value())
    {
        item.version = version.value();
    }
    if (yyjson_val* sha256 = yyjson_obj_get(entry, "sha256Checksum"); sha256 != nullptr && yyjson_is_str(sha256))
    {
        item.sha256Hex.assign(yyjson_get_str(sha256), yyjson_get_len(sha256));
    }
    return S_OK;
}

[[nodiscard]] HRESULT ParseGoogleItemDocument(std::string_view body, FileSystemGoogleDrive::GoogleItem& item)
{
    unique_yyjson_doc doc(yyjson_read(body.data(), body.size(), YYJSON_READ_ALLOW_BOM));
    if (! doc)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    return ParseGoogleItem(yyjson_doc_get_root(doc.get()), item);
}

// Exposed-name resolution shared by every path lookup: the exact Drive name when it is unique among
// its siblings, otherwise the `name [id:...]` decoration. Two children exposing the same text is a
// provider inconsistency that fails closed.
[[nodiscard]] HRESULT ResolveChildByExposedName(const std::vector<FileSystemGoogleDrive::GoogleItem>& children,
                                                std::wstring_view exposedName,
                                                const FileSystemGoogleDrive::GoogleItem*& match)
{
    match = nullptr;
    std::map<std::wstring, size_t, std::less<>> exactNameCounts;
    for (const FileSystemGoogleDrive::GoogleItem& child : children)
    {
        ++exactNameCounts[child.name];
    }
    for (const FileSystemGoogleDrive::GoogleItem& child : children)
    {
        const auto count = exactNameCounts.find(child.name);
        if (MakeExposedItemName(child.name, child.id, count == exactNameCounts.end() ? 0u : count->second) != exposedName)
        {
            continue;
        }
        if (match != nullptr)
        {
            match = nullptr;
            return HRESULT_FROM_WIN32(ERROR_DUP_NAME);
        }
        match = &child;
    }
    return match != nullptr ? S_OK : HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
}

[[nodiscard]] bool IsValidDriveLeafName(std::wstring_view name) noexcept
{
    return ! name.empty() && name.size() <= 255u && name != L"." && name != L".." && name.find(L'/') == std::wstring_view::npos &&
           name.find(L'\0') == std::wstring_view::npos && ! LooksLikeSyntheticDisplayName(name);
}

[[nodiscard]] std::wstring DriveParentPath(std::wstring_view canonicalPath)
{
    const size_t slash = canonicalPath.rfind(L'/');
    return slash == std::wstring_view::npos || slash == 0u ? std::wstring(L"/") : std::wstring(canonicalPath.substr(0, slash));
}

[[nodiscard]] std::wstring_view DriveLeafName(std::wstring_view canonicalPath) noexcept
{
    const size_t slash = canonicalPath.rfind(L'/');
    return slash == std::wstring_view::npos ? canonicalPath : canonicalPath.substr(slash + 1u);
}

[[nodiscard]] bool IsPathWithin(std::wstring_view candidate, std::wstring_view ancestor) noexcept
{
    if (ancestor == L"/")
    {
        return candidate != L"/";
    }
    return candidate.size() > ancestor.size() && candidate.starts_with(ancestor) && candidate[ancestor.size()] == L'/';
}

[[nodiscard]] HRESULT CheckDriveCancel(const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
{
    if (const HRESULT control = FileSystemCheckOperationControl(options); FAILED(control))
    {
        return control;
    }
    if (callback != nullptr)
    {
        BOOL cancel      = FALSE;
        const HRESULT hr = callback->FileSystemShouldCancel(&cancel, cookie);
        if (FAILED(hr))
        {
            return hr;
        }
        if (cancel)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }
    }
    return S_OK;
}

// Mutation receipt: success and definitive refusals are known outcomes; a transport failure,
// cancel, or deadline leaves the outcome unknown (no receipt), which the host reports as such.
[[nodiscard]] const FileSystemItemMutationResult* BuildDriveMutationReceipt(HRESULT itemHr,
                                                                            bool sourceRemovedOnSuccess,
                                                                            FileSystemItemMutationResult& receipt) noexcept
{
    const bool notFound          = IsDriveNotFound(itemHr);
    const bool definitiveRefusal = notFound || itemHr == HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED) || itemHr == HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY) ||
                                   itemHr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) || itemHr == HRESULT_FROM_WIN32(ERROR_INVALID_NAME) ||
                                   itemHr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) || itemHr == HRESULT_FROM_WIN32(ERROR_DUP_NAME) ||
                                   itemHr == HRESULT_FROM_WIN32(ERROR_NOT_SAME_DEVICE) || itemHr == HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER) ||
                                   itemHr == E_INVALIDARG;
    if (FAILED(itemHr) && ! definitiveRefusal)
    {
        return nullptr;
    }
    receipt.sizeBytes             = sizeof(receipt);
    receipt.outcomeKnown          = TRUE;
    receipt.mutationCommitted     = SUCCEEDED(itemHr) ? TRUE : FALSE;
    receipt.originalStillPresent  = notFound ? FALSE : (SUCCEEDED(itemHr) ? (sourceRemovedOnSuccess ? FALSE : TRUE) : TRUE);
    receipt.ownedStageDisposition = FileSystemOwnedStageDisposition::NotApplicable;
    return &receipt;
}

[[nodiscard]] HRESULT ReportDriveItemCompleted(IFileSystemCallback* callback,
                                               FileSystemOperation operation,
                                               unsigned long itemIndex,
                                               const wchar_t* sourcePath,
                                               const wchar_t* destinationPath,
                                               HRESULT status,
                                               bool sourceRemovedOnSuccess,
                                               const FileSystemOptions* options,
                                               void* cookie) noexcept
{
    if (callback == nullptr)
    {
        return S_OK;
    }
    FileSystemItemMutationResult receipt{};
    const FileSystemItemMutationResult* receiptPtr = BuildDriveMutationReceipt(status, sourceRemovedOnSuccess, receipt);
    return callback->FileSystemItemCompleted(
        operation, itemIndex, sourcePath, destinationPath, status, receiptPtr, const_cast<FileSystemOptions*>(options), cookie);
}

void ReportDriveProgress(IFileSystemCallback* callback,
                         FileSystemOperation operation,
                         unsigned long totalItems,
                         unsigned long completedItems,
                         const wchar_t* sourcePath,
                         const wchar_t* destinationPath,
                         const FileSystemOptions* options,
                         void* cookie) noexcept
{
    if (callback != nullptr)
    {
        static_cast<void>(callback->FileSystemProgress(
            operation, totalItems, completedItems, 0u, 0u, sourcePath, destinationPath, 0u, 0u, const_cast<FileSystemOptions*>(options), 0u, cookie));
    }
}

#pragma warning(push)
#pragma warning(disable : 4623 4626 5027) // reference members: no default construction or assignment by design
struct DriveTransferContext final
{
    FileSystemGoogleDrive& fileSystem;
    const FileSystemGoogleDrive::ResolvedConnection& connection;
    FileSystemFlags flags;
    bool move;
    const FileSystemOptions* options;
    IFileSystemCallback* callback;
    void* cookie;

    [[nodiscard]] bool AllowOverwrite() const noexcept
    {
        return (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
    }
    [[nodiscard]] bool ContinueOnError() const noexcept
    {
        return (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
    }
};
#pragma warning(pop)

// A replaced destination file goes to Drive's trash (recoverable) only after the new object exists.
[[nodiscard]] HRESULT TrashReplacedItem(const DriveTransferContext& context, const FileSystemGoogleDrive::GoogleItem& replaced)
{
    FileSystemGoogleDrive::GoogleItem trashed;
    return context.fileSystem.UpdateItem(context.connection, replaced.id, {}, {}, {}, true, trashed);
}

// One file (or shortcut): server-side copy, or a parent/name change for a move. A destination file
// conflict is refused unless overwrite was granted, in which case the old object is trashed after
// the new one exists; a destination folder is never silently replaced by a file.
[[nodiscard]] HRESULT DriveTransferFile(const DriveTransferContext& context,
                                        const FileSystemGoogleDrive::GoogleItem& source,
                                        std::wstring_view sourceParentId,
                                        const FileSystemGoogleDrive::GoogleItem& destinationParent,
                                        std::wstring_view destinationName,
                                        const FileSystemGoogleDrive::GoogleItem* existing)
{
    if (existing != nullptr && (existing->isFolder || ! context.AllowOverwrite()))
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    FileSystemGoogleDrive::GoogleItem result;
    HRESULT hr = S_OK;
    if (context.move)
    {
        const bool sameParent = sourceParentId == destinationParent.id;
        hr                    = context.fileSystem.UpdateItem(context.connection,
                                                              source.id,
                                                              source.name == destinationName ? std::wstring_view{} : destinationName,
                                                              sameParent ? std::wstring_view{} : std::wstring_view(destinationParent.id),
                                                              sameParent ? std::wstring_view{} : sourceParentId,
                                                              std::nullopt,
                                                              result);
    }
    else
    {
        hr = context.fileSystem.CopyFileItem(context.connection, source.id, destinationParent.id, destinationName, result);
    }
    if (FAILED(hr))
    {
        return hr;
    }
    return existing != nullptr ? TrashReplacedItem(context, *existing) : S_OK;
}

// A folder onto a missing name: copy recreates the tree; a move changes the parent natively (the
// caller takes that fast path). A folder onto an existing folder merges per child: colliding files
// are replaced only with overwrite, otherwise skipped, and the item ends as ERROR_PARTIAL_COPY with
// every skipped source preserved.
[[nodiscard]] HRESULT DriveTransferTree(const DriveTransferContext& context,
                                        const FileSystemGoogleDrive::GoogleItem& source,
                                        const FileSystemGoogleDrive::GoogleItem& destinationParent,
                                        std::wstring_view destinationName,
                                        const FileSystemGoogleDrive::GoogleItem* existing)
{
    struct Frame final
    {
        FileSystemGoogleDrive::GoogleItem source;
        FileSystemGoogleDrive::GoogleItem target;
        std::vector<FileSystemGoogleDrive::GoogleItem> children;
        std::vector<FileSystemGoogleDrive::GoogleItem> targetChildren;
        std::map<std::wstring, size_t, std::less<>> exactNameCounts;
        size_t nextChild = 0u;
        bool partial     = false;
    };

    std::vector<Frame> frames;
    const auto pushFrame = [&](const FileSystemGoogleDrive::GoogleItem& frameSource,
                               const FileSystemGoogleDrive::GoogleItem& frameDestinationParent,
                               std::wstring_view frameDestinationName,
                               const FileSystemGoogleDrive::GoogleItem* frameExisting) -> HRESULT
    {
        if (frameExisting != nullptr && ! frameExisting->isFolder)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        Frame frame;
        frame.source = frameSource;
        if (frameExisting != nullptr)
        {
            frame.target = *frameExisting;
        }
        else
        {
            RETURN_IF_FAILED(context.fileSystem.CreateFolderItem(context.connection, frameDestinationParent.id, frameDestinationName, frame.target));
        }
        RETURN_IF_FAILED(context.fileSystem.ListChildren(context.connection, frame.source.id, frame.children));
        if (frameExisting != nullptr)
        {
            RETURN_IF_FAILED(context.fileSystem.ListChildren(context.connection, frame.target.id, frame.targetChildren));
        }
        for (const FileSystemGoogleDrive::GoogleItem& child : frame.children)
        {
            ++frame.exactNameCounts[child.name];
        }
        frames.push_back(std::move(frame));
        return S_OK;
    };

    RETURN_IF_FAILED(pushFrame(source, destinationParent, destinationName, existing));
    while (! frames.empty())
    {
        Frame& frame = frames.back();
        if (frame.nextChild == frame.children.size())
        {
            HRESULT completedHr = frame.partial ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK;
            if (context.move && ! frame.partial)
            {
                // Post-order removal preserves the recursive implementation's rule: a source
                // folder is removed only after every child left it successfully.
                const HRESULT deleteHr = context.fileSystem.DeleteItemPermanently(context.connection, frame.source.id);
                if (FAILED(deleteHr) && ! IsDriveNotFound(deleteHr))
                {
                    completedHr = deleteHr;
                }
            }
            frames.pop_back();
            if (frames.empty())
            {
                return completedHr;
            }
            if (FAILED(completedHr))
            {
                if (completedHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
                {
                    frames.back().partial = true;
                }
                else if (! context.ContinueOnError() || completedHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || completedHr == HRESULT_FROM_WIN32(ERROR_TIMEOUT))
                {
                    return completedHr;
                }
                else
                {
                    frames.back().partial = true;
                }
            }
            continue;
        }

        if (const HRESULT cancel = CheckDriveCancel(context.options, context.callback, context.cookie); FAILED(cancel))
        {
            return cancel;
        }

        const FileSystemGoogleDrive::GoogleItem child = frame.children[frame.nextChild++];
        const auto count                              = frame.exactNameCounts.find(child.name);
        const std::wstring exposedName                = MakeExposedItemName(child.name, child.id, count == frame.exactNameCounts.end() ? 0u : count->second);
        const FileSystemGoogleDrive::GoogleItem* collision = nullptr;
        if (! frame.targetChildren.empty())
        {
            const HRESULT lookup = ResolveChildByExposedName(frame.targetChildren, exposedName, collision);
            if (FAILED(lookup) && ! IsDriveNotFound(lookup))
            {
                return lookup;
            }
        }

        HRESULT childHr = S_OK;
        if (child.isFolder)
        {
            if (context.move && collision == nullptr)
            {
                FileSystemGoogleDrive::GoogleItem moved;
                childHr = context.fileSystem.UpdateItem(context.connection, child.id, {}, frame.target.id, frame.source.id, std::nullopt, moved);
            }
            else
            {
                childHr = pushFrame(child, frame.target, exposedName, collision);
                if (SUCCEEDED(childHr))
                {
                    continue;
                }
            }
        }
        else if (collision != nullptr && (collision->isFolder || ! context.AllowOverwrite()))
        {
            frame.partial = true;
            continue;
        }
        else
        {
            childHr = DriveTransferFile(context, child, frame.source.id, frame.target, exposedName, collision);
        }

        if (FAILED(childHr))
        {
            if (childHr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY))
            {
                frame.partial = true;
                continue;
            }
            if (! context.ContinueOnError() || childHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || childHr == HRESULT_FROM_WIN32(ERROR_TIMEOUT))
            {
                return childHr;
            }
            frame.partial = true;
        }
    }
    return E_UNEXPECTED;
}

[[nodiscard]] HRESULT DriveTransferPath(FileSystemGoogleDrive& fileSystem,
                                        const wchar_t* sourcePath,
                                        const wchar_t* destinationPath,
                                        FileSystemFlags flags,
                                        bool move,
                                        const FileSystemOptions* options,
                                        IFileSystemCallback* callback,
                                        void* cookie)
{
    if (! sourcePath || ! destinationPath || sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    FileSystemGoogleDrive::ResolvedConnection sourceConnection;
    HRESULT hr = fileSystem.ResolveConnection(sourcePath, true, sourceConnection);
    if (FAILED(hr))
    {
        return hr;
    }
    FileSystemGoogleDrive::ResolvedConnection destinationConnection;
    hr = fileSystem.ResolveConnection(destinationPath, true, destinationConnection);
    if (FAILED(hr))
    {
        return hr;
    }
    if (sourceConnection.connectionKey != destinationConnection.connectionKey || sourceConnection.rootKind != destinationConnection.rootKind ||
        sourceConnection.sharedDriveId != destinationConnection.sharedDriveId)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SAME_DEVICE);
    }
    if (sourceConnection.readOnly)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
    if (sourceConnection.canonicalPath == L"/" || destinationConnection.canonicalPath == L"/" ||
        sourceConnection.canonicalPath == destinationConnection.canonicalPath ||
        IsPathWithin(destinationConnection.canonicalPath, sourceConnection.canonicalPath))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER);
    }

    FileSystemGoogleDrive::GoogleItem sourceParent;
    std::wstring sourceLeaf;
    hr = fileSystem.ResolveParentAndLeaf(sourceConnection, sourceConnection.canonicalPath, sourceParent, sourceLeaf);
    if (FAILED(hr))
    {
        return hr;
    }
    FileSystemGoogleDrive::GoogleItem source;
    hr = fileSystem.ResolveChildByName(sourceConnection, sourceParent.id, sourceLeaf, source);
    if (FAILED(hr))
    {
        return hr;
    }

    FileSystemGoogleDrive::GoogleItem destinationParent;
    std::wstring destinationLeaf;
    hr = fileSystem.ResolveParentAndLeaf(destinationConnection, destinationConnection.canonicalPath, destinationParent, destinationLeaf);
    if (FAILED(hr))
    {
        return hr;
    }
    if (! IsValidDriveLeafName(destinationLeaf))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }
    FileSystemGoogleDrive::GoogleItem existing;
    const HRESULT existingHr = fileSystem.ResolveChildByName(destinationConnection, destinationParent.id, destinationLeaf, existing);
    if (FAILED(existingHr) && ! IsDriveNotFound(existingHr))
    {
        return existingHr;
    }
    const FileSystemGoogleDrive::GoogleItem* existingPtr = SUCCEEDED(existingHr) ? &existing : nullptr;
    if (existingPtr != nullptr && existingPtr->id == source.id)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER);
    }

    const DriveTransferContext context{
        .fileSystem = fileSystem,
        .connection = sourceConnection,
        .flags      = flags,
        .move       = move,
        .options    = options,
        .callback   = callback,
        .cookie     = cookie,
    };
    if (! source.isFolder)
    {
        return DriveTransferFile(context, source, sourceParent.id, destinationParent, destinationLeaf, existingPtr);
    }
    if (move && existingPtr == nullptr)
    {
        // Native folder move: one parent/name change re-homes the whole subtree.
        FileSystemGoogleDrive::GoogleItem moved;
        const bool sameParent = sourceParent.id == destinationParent.id;
        return fileSystem.UpdateItem(sourceConnection,
                                     source.id,
                                     source.name == destinationLeaf ? std::wstring_view{} : std::wstring_view(destinationLeaf),
                                     sameParent ? std::wstring_view{} : std::wstring_view(destinationParent.id),
                                     sameParent ? std::wstring_view{} : std::wstring_view(sourceParent.id),
                                     std::nullopt,
                                     moved);
    }
    return DriveTransferTree(context, source, destinationParent, destinationLeaf, existingPtr);
}

// Delete: permanent DELETE or trash (recycle). A folder without RECURSIVE must be empty; with
// RECURSIVE (or when trashing) Drive removes the subtree as one object.
[[nodiscard]] HRESULT DriveDeletePath(FileSystemGoogleDrive& fileSystem, const wchar_t* path, FileSystemFlags flags)
{
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }
    FileSystemGoogleDrive::ResolvedConnection connection;
    HRESULT hr = fileSystem.ResolveConnection(path, true, connection);
    if (FAILED(hr))
    {
        return hr;
    }
    if (connection.readOnly)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
    if (connection.canonicalPath == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
    FileSystemGoogleDrive::GoogleItem item;
    hr = fileSystem.ResolveItemByPath(connection, connection.canonicalPath, item);
    if (FAILED(hr))
    {
        return hr;
    }
    if (item.isFolder && (flags & FILESYSTEM_FLAG_RECURSIVE) == 0)
    {
        std::vector<FileSystemGoogleDrive::GoogleItem> children;
        hr = fileSystem.ListChildren(connection, item.id, children);
        if (FAILED(hr))
        {
            return hr;
        }
        if (! children.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
        }
    }
    if ((flags & FILESYSTEM_FLAG_USE_RECYCLE_BIN) != 0)
    {
        FileSystemGoogleDrive::GoogleItem trashed;
        return fileSystem.UpdateItem(connection, item.id, {}, {}, {}, true, trashed);
    }
    return fileSystem.DeleteItemPermanently(connection, item.id);
}

// C10: resolve the item the path names now and report its Drive file id as the delete identity.
[[nodiscard]] HRESULT DriveResolveDeleteIdentity(FileSystemGoogleDrive& fileSystem, const wchar_t* path, FileSystemDeleteIdentity& identity)
{
    identity.isDirectory = FALSE;
    identity.identity[0] = L'\0';
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }
    FileSystemGoogleDrive::ResolvedConnection connection;
    HRESULT hr = fileSystem.ResolveConnection(path, true, connection);
    if (FAILED(hr))
    {
        return hr;
    }
    if (connection.canonicalPath == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
    FileSystemGoogleDrive::GoogleItem item;
    hr = fileSystem.ResolveItemByPath(connection, connection.canonicalPath, item);
    if (FAILED(hr))
    {
        return hr;
    }
    if (item.id.empty() || item.id.size() >= std::size(identity.identity))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    identity.isDirectory = item.isFolder ? TRUE : FALSE;
    std::copy(item.id.begin(), item.id.end(), identity.identity);
    identity.identity[item.id.size()] = L'\0';
    return S_OK;
}

// C10: DriveDeletePath, deleting only while the path still names the pinned file id.
[[nodiscard]] HRESULT DriveDeletePathIfIdentity(FileSystemGoogleDrive& fileSystem,
                                                const wchar_t* path,
                                                const FileSystemDeleteIdentity& pinned,
                                                FileSystemFlags flags)
{
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }
    FileSystemGoogleDrive::ResolvedConnection connection;
    HRESULT hr = fileSystem.ResolveConnection(path, true, connection);
    if (FAILED(hr))
    {
        return hr;
    }
    if (connection.readOnly)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
    if (connection.canonicalPath == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
    FileSystemGoogleDrive::GoogleItem item;
    hr = fileSystem.ResolveItemByPath(connection, connection.canonicalPath, item);
    if (FAILED(hr))
    {
        return hr;
    }
    if (std::wstring_view(pinned.identity) != std::wstring_view(item.id) || (item.isFolder ? TRUE : FALSE) != pinned.isDirectory)
    {
        // The name refers to another object than the one the user confirmed; nothing is deleted.
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
    if (item.isFolder && (flags & FILESYSTEM_FLAG_RECURSIVE) == 0)
    {
        std::vector<FileSystemGoogleDrive::GoogleItem> children;
        hr = fileSystem.ListChildren(connection, item.id, children);
        if (FAILED(hr))
        {
            return hr;
        }
        if (! children.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
        }
    }
    return fileSystem.DeleteItemPermanently(connection, item.id);
}

[[nodiscard]] std::wstring JoinDrivePath(std::wstring_view folder, std::wstring_view leaf)
{
    std::wstring joined(folder);
    if (joined.empty() || joined.back() != L'/')
    {
        joined.push_back(L'/');
    }
    joined.append(leaf);
    return joined;
}
} // namespace

HRESULT FileSystemGoogleDrive::GetItemById(const ResolvedConnection& connection, std::wstring_view id, GoogleItem& item)
{
    item = {};
    if (id.empty())
    {
        return E_INVALIDARG;
    }
    std::string url = FilesEndpointUrl() + "/" + UrlEncodeUtf8(Utf8FromUtf16(id));
    AppendQueryParam(url, "supportsAllDrives", "true");
    AppendQueryParam(url, "fields", std::string(kDriveItemFields));
    std::string body;
    const HRESULT hr = PerformAuthorizedJsonGet(connection, url, body);
    if (FAILED(hr))
    {
        return hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ? HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) : hr;
    }
    return ParseGoogleItemDocument(body, item);
}

HRESULT FileSystemGoogleDrive::ResolveChildByName(const ResolvedConnection& connection,
                                                  std::wstring_view parentId,
                                                  std::wstring_view exposedName,
                                                  GoogleItem& child)
{
    child = {};
    std::vector<GoogleItem> children;
    const HRESULT hr = ListChildren(connection, parentId, children);
    if (FAILED(hr))
    {
        return hr;
    }
    const GoogleItem* match = nullptr;
    const HRESULT lookup    = ResolveChildByExposedName(children, exposedName, match);
    if (FAILED(lookup))
    {
        return lookup;
    }
    child = *match;
    return S_OK;
}

HRESULT FileSystemGoogleDrive::ResolveParentAndLeaf(const ResolvedConnection& connection,
                                                    std::wstring_view canonicalPath,
                                                    GoogleItem& parent,
                                                    std::wstring& leafName)
{
    parent = {};
    leafName.clear();
    const std::wstring normalized = NormalizePluginPath(canonicalPath);
    if (normalized == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER);
    }
    leafName.assign(DriveLeafName(normalized));
    const HRESULT hr = ResolveItemByPath(connection, DriveParentPath(normalized), parent);
    if (FAILED(hr))
    {
        return hr;
    }
    return parent.isFolder ? S_OK : HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
}

HRESULT FileSystemGoogleDrive::CreateFolderItem(const ResolvedConnection& connection, std::wstring_view parentId, std::wstring_view name, GoogleItem& created)
{
    created         = {};
    std::string url = FilesEndpointUrl();
    AppendQueryParam(url, "supportsAllDrives", "true");
    AppendQueryParam(url, "fields", std::string(kDriveItemFields));
    const std::string body = std::format(R"({{"name":{},"mimeType":"{}","parents":[{}]}})", JsonQuote(name), kFolderMimeType, JsonQuote(parentId));
    const std::vector<std::string> headers{"Content-Type: application/json; charset=UTF-8"};
    long statusCode = 0;
    std::string responseBody;
    const HRESULT hr = PerformAuthorizedRequest(
        connection, "POST", url, headers, body, kMaxJsonResponseBytes, statusCode, responseBody, nullptr, nullptr, AuthorizedRequestRetry::Mutation);
    if (FAILED(hr))
    {
        return hr;
    }
    if (const HRESULT statusHr = MapHttpStatusToHresult(statusCode); FAILED(statusHr))
    {
        return statusHr;
    }
    return ParseGoogleItemDocument(responseBody, created);
}

HRESULT FileSystemGoogleDrive::UpdateItem(const ResolvedConnection& connection,
                                          std::wstring_view id,
                                          std::wstring_view newName,
                                          std::wstring_view addParent,
                                          std::wstring_view removeParent,
                                          std::optional<bool> trashed,
                                          GoogleItem& updated)
{
    updated = {};
    if (id.empty())
    {
        return E_INVALIDARG;
    }
    std::string url = FilesEndpointUrl() + "/" + UrlEncodeUtf8(Utf8FromUtf16(id));
    AppendQueryParam(url, "supportsAllDrives", "true");
    AppendQueryParam(url, "fields", std::string(kDriveItemFields));
    if (! addParent.empty())
    {
        AppendQueryParam(url, "addParents", Utf8FromUtf16(addParent));
    }
    if (! removeParent.empty())
    {
        AppendQueryParam(url, "removeParents", Utf8FromUtf16(removeParent));
    }
    std::string body = "{";
    if (! newName.empty())
    {
        body += std::format(R"("name":{})", JsonQuote(newName));
    }
    if (trashed.has_value())
    {
        if (body.size() > 1u)
        {
            body += ",";
        }
        body += std::format(R"("trashed":{})", trashed.value() ? "true" : "false");
    }
    body += "}";
    const std::vector<std::string> headers{"Content-Type: application/json; charset=UTF-8"};
    long statusCode = 0;
    std::string responseBody;
    const HRESULT hr = PerformAuthorizedRequest(
        connection, "PATCH", url, headers, body, kMaxJsonResponseBytes, statusCode, responseBody, nullptr, nullptr, AuthorizedRequestRetry::Mutation);
    if (FAILED(hr))
    {
        return hr;
    }
    if (const HRESULT statusHr = MapHttpStatusToHresult(statusCode); FAILED(statusHr))
    {
        return statusHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ? HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) : statusHr;
    }
    // A trash request reports the (now trashed) object; parse it as a plain resource.
    unique_yyjson_doc doc(yyjson_read(responseBody.data(), responseBody.size(), YYJSON_READ_ALLOW_BOM));
    if (! doc)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    const HRESULT parseHr = ParseGoogleItem(yyjson_doc_get_root(doc.get()), updated);
    if (parseHr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) && trashed.value_or(false))
    {
        return S_OK;
    }
    return parseHr;
}

HRESULT FileSystemGoogleDrive::DeleteItemPermanently(const ResolvedConnection& connection, std::wstring_view id)
{
    if (id.empty())
    {
        return E_INVALIDARG;
    }
    std::string url = FilesEndpointUrl() + "/" + UrlEncodeUtf8(Utf8FromUtf16(id));
    AppendQueryParam(url, "supportsAllDrives", "true");
    long statusCode = 0;
    std::string responseBody;
    const HRESULT hr = PerformAuthorizedRequest(
        connection, "DELETE", url, {}, {}, kMaxJsonResponseBytes, statusCode, responseBody, nullptr, nullptr, AuthorizedRequestRetry::Mutation);
    if (FAILED(hr))
    {
        return hr;
    }
    const HRESULT statusHr = MapHttpStatusToHresult(statusCode);
    return statusHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ? HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) : statusHr;
}

HRESULT FileSystemGoogleDrive::CopyFileItem(
    const ResolvedConnection& connection, std::wstring_view id, std::wstring_view parentId, std::wstring_view name, GoogleItem& copy)
{
    copy = {};
    if (id.empty() || parentId.empty())
    {
        return E_INVALIDARG;
    }
    std::string url = FilesEndpointUrl() + "/" + UrlEncodeUtf8(Utf8FromUtf16(id)) + "/copy";
    AppendQueryParam(url, "supportsAllDrives", "true");
    AppendQueryParam(url, "fields", std::string(kDriveItemFields));
    const std::string body = std::format(R"({{"name":{},"parents":[{}]}})", JsonQuote(name), JsonQuote(parentId));
    const std::vector<std::string> headers{"Content-Type: application/json; charset=UTF-8"};
    long statusCode = 0;
    std::string responseBody;
    const HRESULT hr = PerformAuthorizedRequest(
        connection, "POST", url, headers, body, kMaxJsonResponseBytes, statusCode, responseBody, nullptr, nullptr, AuthorizedRequestRetry::Mutation);
    if (FAILED(hr))
    {
        return hr;
    }
    if (const HRESULT statusHr = MapHttpStatusToHresult(statusCode); FAILED(statusHr))
    {
        return statusHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ? HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) : statusHr;
    }
    return ParseGoogleItemDocument(responseBody, copy);
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::CopyItem(const wchar_t* sourcePath,
                                                          const wchar_t* destinationPath,
                                                          FileSystemFlags flags,
                                                          const FileSystemOptions* options,
                                                          IFileSystemCallback* callback,
                                                          void* cookie) noexcept
{
    const DriveOperationOptionsScope operationOptionsScope(options);
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    try
    {
        const HRESULT hr         = DriveTransferPath(*this, sourcePath, destinationPath, flags, false, options, callback, cookie);
        const HRESULT callbackHr = ReportDriveItemCompleted(callback, FILESYSTEM_COPY, 0u, sourcePath, destinationPath, hr, false, options, cookie);
        return FAILED(callbackHr) ? callbackHr : hr;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive CopyItem failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::MoveItem(const wchar_t* sourcePath,
                                                          const wchar_t* destinationPath,
                                                          FileSystemFlags flags,
                                                          const FileSystemOptions* options,
                                                          IFileSystemCallback* callback,
                                                          void* cookie) noexcept
{
    const DriveOperationOptionsScope operationOptionsScope(options);
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    try
    {
        const HRESULT hr         = DriveTransferPath(*this, sourcePath, destinationPath, flags, true, options, callback, cookie);
        const HRESULT callbackHr = ReportDriveItemCompleted(callback, FILESYSTEM_MOVE, 0u, sourcePath, destinationPath, hr, true, options, cookie);
        return FAILED(callbackHr) ? callbackHr : hr;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive MoveItem failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::DeleteItem(
    const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
{
    const DriveOperationOptionsScope operationOptionsScope(options);
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    try
    {
        const HRESULT hr         = DriveDeletePath(*this, path, flags);
        const HRESULT callbackHr = ReportDriveItemCompleted(callback, FILESYSTEM_DELETE, 0u, path, nullptr, hr, true, options, cookie);
        return FAILED(callbackHr) ? callbackHr : hr;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive DeleteItem failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::ResolveDeleteIdentity(const wchar_t* path,
                                                                       const FileSystemOptions* options,
                                                                       FileSystemDeleteIdentity* identity) noexcept
{
    const DriveOperationOptionsScope operationOptionsScope(options);
    if (path == nullptr || identity == nullptr)
    {
        return E_POINTER;
    }
    if (identity->sizeBytes < sizeof(FileSystemDeleteIdentity))
    {
        return E_INVALIDARG;
    }
    try
    {
        return DriveResolveDeleteIdentity(*this, path, *identity);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive ResolveDeleteIdentity failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::DeleteIfIdentity(const wchar_t* path,
                                                                  const FileSystemDeleteIdentity* identity,
                                                                  FileSystemFlags flags,
                                                                  const FileSystemOptions* options,
                                                                  IFileSystemCallback* callback,
                                                                  void* cookie) noexcept
{
    const DriveOperationOptionsScope operationOptionsScope(options);
    if (path == nullptr || identity == nullptr)
    {
        return E_POINTER;
    }
    if (identity->sizeBytes < sizeof(FileSystemDeleteIdentity) || identity->identity[0] == L'\0' || ! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    try
    {
        const HRESULT hr         = DriveDeletePathIfIdentity(*this, path, *identity, flags);
        const HRESULT callbackHr = ReportDriveItemCompleted(callback, FILESYSTEM_DELETE, 0u, path, nullptr, hr, true, options, cookie);
        return FAILED(callbackHr) ? callbackHr : hr;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive DeleteIfIdentity failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::RenameItem(const wchar_t* sourcePath,
                                                            const wchar_t* destinationPath,
                                                            FileSystemFlags flags,
                                                            const FileSystemOptions* options,
                                                            IFileSystemCallback* callback,
                                                            void* cookie) noexcept
{
    const DriveOperationOptionsScope operationOptionsScope(options);
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    try
    {
        // A rename is a move whose destination spells the new name; Drive applies both as one PATCH.
        const HRESULT hr         = DriveTransferPath(*this, sourcePath, destinationPath, flags, true, options, callback, cookie);
        const HRESULT callbackHr = ReportDriveItemCompleted(callback, FILESYSTEM_RENAME, 0u, sourcePath, destinationPath, hr, true, options, cookie);
        return FAILED(callbackHr) ? callbackHr : hr;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive RenameItem failed with std::exception.");
        return E_FAIL;
    }
}

namespace
{
// Batch entry points share one loop: per item cancel check, the mutation, a completion with its
// receipt, and CONTINUE_ON_ERROR semantics (first failure wins the batch status).
template <typename ItemFn>
[[nodiscard]] HRESULT RunDriveBatch(
    unsigned long count, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie, ItemFn&& itemFn)
{
    const bool continueOnError = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
    HRESULT firstFailure       = S_OK;
    for (unsigned long index = 0; index < count; ++index)
    {
        if (const HRESULT cancel = CheckDriveCancel(options, callback, cookie); FAILED(cancel))
        {
            return cancel;
        }
        const HRESULT itemHr = itemFn(index);
        if (FAILED(itemHr))
        {
            if (itemHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || itemHr == HRESULT_FROM_WIN32(ERROR_TIMEOUT) || itemHr == E_OUTOFMEMORY)
            {
                return itemHr;
            }
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = itemHr;
            }
            if (! continueOnError)
            {
                return itemHr;
            }
        }
    }
    return firstFailure;
}
} // namespace

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::CopyItems(const wchar_t* const* sourcePaths,
                                                           unsigned long count,
                                                           const wchar_t* destinationFolder,
                                                           FileSystemFlags flags,
                                                           const FileSystemOptions* options,
                                                           IFileSystemCallback* callback,
                                                           void* cookie) noexcept
{
    const DriveOperationOptionsScope operationOptionsScope(options);
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    if ((! sourcePaths && count != 0) || ! destinationFolder)
    {
        return E_POINTER;
    }
    try
    {
        return RunDriveBatch(count,
                             flags,
                             options,
                             callback,
                             cookie,
                             [&](unsigned long index)
        {
            const wchar_t* sourcePath = sourcePaths[index];
            std::wstring destination;
            HRESULT hr = sourcePath && sourcePath[0] != L'\0' ? S_OK : E_INVALIDARG;
            if (SUCCEEDED(hr))
            {
                destination = JoinDrivePath(destinationFolder, DriveLeafName(NormalizePluginPath(sourcePath)));
                hr          = DriveTransferPath(*this, sourcePath, destination.c_str(), flags, false, options, callback, cookie);
            }
            const HRESULT callbackHr = ReportDriveItemCompleted(callback, FILESYSTEM_COPY, index, sourcePath, destination.c_str(), hr, false, options, cookie);
            ReportDriveProgress(callback, FILESYSTEM_COPY, count, index + 1u, sourcePath, destination.c_str(), options, cookie);
            return FAILED(callbackHr) ? callbackHr : hr;
        });
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive CopyItems failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::MoveItems(const wchar_t* const* sourcePaths,
                                                           unsigned long count,
                                                           const wchar_t* destinationFolder,
                                                           FileSystemFlags flags,
                                                           const FileSystemOptions* options,
                                                           IFileSystemCallback* callback,
                                                           void* cookie) noexcept
{
    const DriveOperationOptionsScope operationOptionsScope(options);
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    if ((! sourcePaths && count != 0) || ! destinationFolder)
    {
        return E_POINTER;
    }
    try
    {
        return RunDriveBatch(count,
                             flags,
                             options,
                             callback,
                             cookie,
                             [&](unsigned long index)
        {
            const wchar_t* sourcePath = sourcePaths[index];
            std::wstring destination;
            HRESULT hr = sourcePath && sourcePath[0] != L'\0' ? S_OK : E_INVALIDARG;
            if (SUCCEEDED(hr))
            {
                destination = JoinDrivePath(destinationFolder, DriveLeafName(NormalizePluginPath(sourcePath)));
                hr          = DriveTransferPath(*this, sourcePath, destination.c_str(), flags, true, options, callback, cookie);
            }
            const HRESULT callbackHr = ReportDriveItemCompleted(callback, FILESYSTEM_MOVE, index, sourcePath, destination.c_str(), hr, true, options, cookie);
            ReportDriveProgress(callback, FILESYSTEM_MOVE, count, index + 1u, sourcePath, destination.c_str(), options, cookie);
            return FAILED(callbackHr) ? callbackHr : hr;
        });
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive MoveItems failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::DeleteItems(const wchar_t* const* paths,
                                                             unsigned long count,
                                                             FileSystemFlags flags,
                                                             const FileSystemOptions* options,
                                                             IFileSystemCallback* callback,
                                                             void* cookie) noexcept
{
    const DriveOperationOptionsScope operationOptionsScope(options);
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    if (! paths && count != 0)
    {
        return E_POINTER;
    }
    try
    {
        return RunDriveBatch(count,
                             flags,
                             options,
                             callback,
                             cookie,
                             [&](unsigned long index)
        {
            const wchar_t* path      = paths[index];
            const HRESULT hr         = DriveDeletePath(*this, path, flags);
            const HRESULT callbackHr = ReportDriveItemCompleted(callback, FILESYSTEM_DELETE, index, path, nullptr, hr, true, options, cookie);
            ReportDriveProgress(callback, FILESYSTEM_DELETE, count, index + 1u, path, nullptr, options, cookie);
            return FAILED(callbackHr) ? callbackHr : hr;
        });
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive DeleteItems failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::RenameItems(const FileSystemRenamePair* items,
                                                             unsigned long count,
                                                             FileSystemFlags flags,
                                                             const FileSystemOptions* options,
                                                             IFileSystemCallback* callback,
                                                             void* cookie) noexcept
{
    const DriveOperationOptionsScope operationOptionsScope(options);
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    if (! items && count != 0)
    {
        return E_POINTER;
    }
    try
    {
        return RunDriveBatch(count,
                             flags,
                             options,
                             callback,
                             cookie,
                             [&](unsigned long index)
        {
            const FileSystemRenamePair& pair = items[index];
            std::wstring destination;
            HRESULT hr = pair.sizeBytes >= sizeof(FileSystemRenamePair) && pair.sourcePath && pair.sourcePath[0] != L'\0' && pair.newName ? S_OK : E_INVALIDARG;
            if (SUCCEEDED(hr))
            {
                // The pair carries the new leaf name; the destination keeps the source's parent.
                destination = JoinDrivePath(DriveParentPath(NormalizePluginPath(pair.sourcePath)), pair.newName);
                hr          = DriveTransferPath(*this, pair.sourcePath, destination.c_str(), flags, true, options, callback, cookie);
            }
            const HRESULT callbackHr =
                ReportDriveItemCompleted(callback, FILESYSTEM_RENAME, index, pair.sourcePath, destination.c_str(), hr, true, options, cookie);
            ReportDriveProgress(callback, FILESYSTEM_RENAME, count, index + 1u, pair.sourcePath, destination.c_str(), options, cookie);
            return FAILED(callbackHr) ? callbackHr : hr;
        });
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive RenameItems failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetPathCapabilities(const wchar_t* path, FileSystemOperation operation, const char** jsonUtf8) noexcept
{
    if (! jsonUtf8)
    {
        return E_POINTER;
    }
    *jsonUtf8 = nullptr;
    if (! path || path[0] == L'\0' || operation < FILESYSTEM_COPY || operation > FILESYSTEM_CREATE_DIRECTORY)
    {
        return E_INVALIDARG;
    }

    std::wstring rootIdentity = L"provider-root";
    if (NormalizePluginPath(path) != L"/")
    {
        ResolvedConnection connection;
        const HRESULT resolveHr = ResolveConnection(path, false, connection);
        if (FAILED(resolveHr))
        {
            return resolveHr;
        }
        rootIdentity = connection.rootKind == L"sharedDrive" && ! connection.sharedDriveId.empty() ? connection.sharedDriveId : connection.connectionName;
    }

    std::string encodedRoot;
    if (! Common::Uri::TryPercentEncodeUtf8(rootIdentity, Common::Uri::SlashPolicy::Encode, encodedRoot) || encodedRoot.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    std::lock_guard lock(_stateMutex);
    _capabilitiesJson.assign(kCapabilitiesJson);
    constexpr std::string_view placeholder = "configured-drive-root";
    const size_t offset                    = _capabilitiesJson.find(placeholder);
    if (offset == std::string::npos)
    {
        _capabilitiesJson.clear();
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    _capabilitiesJson.replace(offset, placeholder.size(), std::format("google-drive:{}", encodedRoot));
    constexpr std::string_view watchdogPlaceholder = "configured-drive-watchdog-ms";
    const size_t watchdogOffset                    = _capabilitiesJson.find(watchdogPlaceholder);
    if (watchdogOffset == std::string::npos)
    {
        _capabilitiesJson.clear();
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    _capabilitiesJson.replace(
        watchdogOffset,
        watchdogPlaceholder.size(),
        std::to_string(FileSystemGoogleDriveInternal::DriveProviderWatchdogTimeoutMs(_settings.connectTimeoutMs, _settings.requestTimeoutMs)));
    *jsonUtf8 = _capabilitiesJson.c_str();
    return S_OK;
}

HRESULT FileSystemGoogleDrive::BuildFileSystemRouteDescriptor(const wchar_t* path,
                                                              FileSystemOperation operation,
                                                              FileSystemRouteDescriptor& descriptor) noexcept
{
    static_cast<void>(operation);
    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    std::wstring rootIdentity = L"provider-root";
    bool readOnly             = false; // a read-only connection profile keeps the route read-only
    if (NormalizePluginPath(path) != L"/")
    {
        ResolvedConnection connection;
        const HRESULT resolveHr = ResolveConnection(path, false, connection);
        if (FAILED(resolveHr))
        {
            return resolveHr;
        }
        rootIdentity = connection.rootKind == L"sharedDrive" && ! connection.sharedDriveId.empty() ? connection.sharedDriveId : connection.connectionName;
        readOnly     = connection.readOnly;
    }

    std::string encodedRoot;
    if (! Common::Uri::TryPercentEncodeUtf8(rootIdentity, Common::Uri::SlashPolicy::Encode, encodedRoot) || encodedRoot.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    Settings settings{};
    {
        std::scoped_lock lock(_stateMutex);
        settings = _settings;
    }

    descriptor               = {};
    descriptor.providerId    = kPluginId;
    descriptor.pathProfileId = L"google-drive";
    descriptor.rootId        = std::format(L"google-drive:{}", Utf16FromUtf8(encodedRoot));
    descriptor.availability  = FILESYSTEM_ROUTE_AVAILABLE;
    // R0f-GDrive: every libcurl transfer under a task polls the operation control from its progress
    // callback and is bounded by the connect and hard request timeouts.
    descriptor.cancellationRoute         = FILESYSTEM_CANCELLATION_PROVIDER_WATCHDOG;
    descriptor.cancellationDeadline      = true;
    descriptor.providerWatchdogTimeoutMs = FileSystemGoogleDriveInternal::DriveProviderWatchdogTimeoutMs(settings.connectTimeoutMs, settings.requestTimeoutMs);
    descriptor.namespaceKind             = FILESYSTEM_NAMESPACE_PROVIDER_VIRTUAL_FOLDER;
    descriptor.componentComparison       = FILESYSTEM_ROUTE_COMPONENT_ORDINAL_CASE_SENSITIVE;
    descriptor.caseOnlyRename            = FILESYSTEM_ROUTE_CASE_ONLY_SUPPORTED;
    // Exposed names are unique per folder (duplicate Drive names carry the `[id:...]` decoration) and
    // an ambiguous text fails closed with ERROR_DUP_NAME, so the exposed path text names one object.
    descriptor.pathTextStableIdentity         = true;
    descriptor.proofFlags                     = FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST; // R3-2: sha256Checksum of the published object
    descriptor.copyMoveMaxConcurrency         = 4u;
    descriptor.deleteMaxConcurrency           = 4u;
    descriptor.deleteRecycleBinMaxConcurrency = 1u;
    descriptor.maxComponentUtf16              = 255u;
    const bool writable                       = ! readOnly;
    descriptor.copyOperation                  = writable;
    descriptor.moveOperation                  = writable;
    descriptor.nativeMoveOperation            = writable;
    descriptor.deleteOperation                = writable;
    descriptor.renameOperation                = writable;
    descriptor.createDirectoryOperation       = writable;
    descriptor.propertiesOperation            = true;
    descriptor.readOperation                  = true;
    descriptor.writeOperation                 = writable;
    descriptor.recycleOperation               = writable;
    descriptor.committedSize                  = true;
    descriptor.exportCopyAll                  = true;
    descriptor.importCopyAll                  = writable;
    return descriptor.rootId.empty() ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA) : S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetTransferHints([[maybe_unused]] const wchar_t* path,
                                                                  [[maybe_unused]] FileSystemOperation operationType,
                                                                  [[maybe_unused]] FileSystemTransferEndpoint endpoint,
                                                                  FileSystemTransferHints* hints) noexcept
{
    if (! path || path[0] == L'\0' || ! hints)
    {
        return E_INVALIDARG;
    }
    if (hints->sizeBytes < sizeof(FileSystemTransferHints))
    {
        return E_INVALIDARG;
    }

    hints->latencyClass = FILESYSTEM_TRANSFER_LATENCY_CLOUD;
    hints->flags =
        FILESYSTEM_TRANSFER_HINT_PREFERS_LARGE_BUFFERS | FILESYSTEM_TRANSFER_HINT_PREFERS_SEQUENTIAL_IO | FILESYSTEM_TRANSFER_HINT_HIGH_METADATA_COST;
    hints->preferredBufferBytes      = 8u * 1024u * 1024u;
    hints->preferredProgressPeriodMs = 200u;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetStorageCharacteristics([[maybe_unused]] const wchar_t* path,
                                                                           FileSystemStorageCharacteristics* characteristics) noexcept
{
    if (! path || path[0] == L'\0' || ! characteristics)
    {
        return E_INVALIDARG;
    }
    if (characteristics->sizeBytes < sizeof(FileSystemStorageCharacteristics))
    {
        return E_INVALIDARG;
    }

    characteristics->storageKind = FILESYSTEM_STORAGE_CLOUD;
    characteristics->flags = FILESYSTEM_STORAGE_FLAG_HIGH_LATENCY | FILESYSTEM_STORAGE_FLAG_PREFERS_SEQUENTIAL_IO | FILESYSTEM_STORAGE_FLAG_SUPPORTS_DEEP_QUEUE;
    characteristics->queueDepthHint               = 8u;
    characteristics->preferredCopyMoveConcurrency = 8u;
    characteristics->preferredDeleteConcurrency   = 8u;
    return S_OK;
}

HRESULT FileSystemGoogleDrive::SetConfigurationImpl(const char* configurationJsonUtf8)
{
    Settings newSettings{};
    std::string configuration = "{}";

    if (configurationJsonUtf8 && configurationJsonUtf8[0] != '\0')
    {
        configuration                       = configurationJsonUtf8;
        Common::Json::ObjectDocument parsed = Common::Json::ParseObjectDocument(configuration, YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (! parsed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        yyjson_val* root = parsed.root;
        if (const auto value = TryGetJsonString(root, "defaultClientId"); value.has_value())
        {
            newSettings.defaultClientId = value.value();
        }
        if (const auto value = TryGetJsonUInt(root, "connectTimeoutMs"); value.has_value() && value.value() >= 1u)
        {
            newSettings.connectTimeoutMs = static_cast<uint32_t>((std::min)(value.value(), 600'000ull));
        }
        if (const auto value = TryGetJsonUInt(root, "requestTimeoutMs"); value.has_value() && value.value() >= 1u)
        {
            newSettings.requestTimeoutMs = static_cast<uint32_t>((std::min)(value.value(), 600'000ull));
        }
        if (const auto value = TryGetJsonUInt(root, "pageSize"); value.has_value() && value.value() >= 1u)
        {
            newSettings.pageSize = static_cast<unsigned long>((std::min)(value.value(), 1000ull));
        }
    }

    {
        std::scoped_lock lock(_stateMutex);
        _settings = std::move(newSettings);
        // Write to the inactive buffer and then flip the index atomically
        const size_t nextIndex               = 1 - _configurationJsonIndex;
        _configurationJsonStorage[nextIndex] = std::move(configuration);
        _configurationJsonIndex              = nextIndex;
    }

    {
        std::scoped_lock lock(_tokenMutex);
        _accessTokensByConnectionKey.clear();
        _lastTokenRefreshStatus.clear();
        ++_tokenCacheGeneration;
    }

    return S_OK;
}

HRESULT FileSystemGoogleDrive::ResolveConnection(const wchar_t* path, bool acquireSecrets, ResolvedConnection& outConnection)
{
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

#if defined(ENABLE_TESTS)
    if (g_debugDriveSyntheticConnection.load(std::memory_order_acquire))
    {
        // The loopback fixture has no connection profile: `/@conn:google-drive-selftest/...` resolves
        // to a synthetic My Drive connection whose token the fixture's token endpoint answers.
        constexpr std::wstring_view kDebugPrefix = L"/@conn:google-drive-selftest";
        const std::wstring canonical             = NormalizePluginPath(path);
        if (! OrdinalString::StartsWithNoCase(canonical, kDebugPrefix) || (canonical.size() > kDebugPrefix.size() && canonical[kDebugPrefix.size()] != L'/'))
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_CONNECTED);
        }
        Settings defaults{};
        {
            std::scoped_lock lock(_stateMutex);
            defaults = _settings;
        }
        outConnection                = {};
        outConnection.connectionName = L"google-drive-selftest";
        outConnection.canonicalPath = canonical.size() == kDebugPrefix.size() ? std::wstring(L"/") : NormalizePluginPath(canonical.substr(kDebugPrefix.size()));
        outConnection.clientId      = L"debug-client";
        outConnection.connectionKey = ConnectionKeyForCache(outConnection.connectionName, outConnection.clientId);
        outConnection.rootKind      = L"myDrive";
        outConnection.pageSize      = defaults.pageSize;
        outConnection.connectTimeoutMs = defaults.connectTimeoutMs;
        outConnection.requestTimeoutMs = defaults.requestTimeoutMs;
        if (acquireSecrets)
        {
            outConnection.refreshToken = L"debug-refresh-token";
        }
        return S_OK;
    }
#endif

    Settings defaults{};
    wil::com_ptr<IHostConnections> hostConnections;
    {
        std::scoped_lock lock(_stateMutex);
        defaults        = _settings;
        hostConnections = _hostConnections;
    }

    if (! hostConnections)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const std::wstring normalizedFull = NormalizePluginPath(path);

    std::wstring_view authority;
    std::wstring_view pathPart;
    if (normalizedFull.size() >= 2u && normalizedFull[0] == L'/' && normalizedFull[1] == L'/')
    {
        std::wstring_view after(normalizedFull);
        after.remove_prefix(2);

        const size_t slashPos = after.find(L'/');
        authority             = (slashPos == std::wstring_view::npos) ? after : after.substr(0, slashPos);
        pathPart              = (slashPos == std::wstring_view::npos) ? std::wstring_view(L"/") : after.substr(slashPos);
    }
    else
    {
        pathPart = normalizedFull;
    }

    bool hasConnectionPrefix = false;
    std::wstring_view connectionName;
    std::wstring_view connectionPath = pathPart;

    std::wstring_view rest = pathPart;
    while (! rest.empty() && rest.front() == L'/')
    {
        rest.remove_prefix(1);
    }

    constexpr std::wstring_view kConnectionPrefix = L"@conn:";
    if (OrdinalString::StartsWithNoCase(rest, kConnectionPrefix))
    {
        rest.remove_prefix(kConnectionPrefix.size());
        const size_t slashPos = rest.find(L'/');
        connectionName        = (slashPos == std::wstring_view::npos) ? rest : rest.substr(0, slashPos);
        connectionPath        = (slashPos == std::wstring_view::npos) ? std::wstring_view(L"/") : rest.substr(slashPos);
        hasConnectionPrefix   = true;
    }
    else if (OrdinalString::EqualsNoCase(authority, L"@conn"))
    {
        const size_t slashPos = rest.find(L'/');
        connectionName        = (slashPos == std::wstring_view::npos) ? rest : rest.substr(0, slashPos);
        connectionPath        = (slashPos == std::wstring_view::npos) ? std::wstring_view(L"/") : rest.substr(slashPos);
        hasConnectionPrefix   = true;
    }

    if (! hasConnectionPrefix || connectionName.empty())
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

    if (! json || json.get()[0] == '\0')
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    unique_yyjson_doc doc(yyjson_read(json.get(), std::strlen(json.get()), YYJSON_READ_ALLOW_BOM));
    if (! doc)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    yyjson_val* root = yyjson_doc_get_root(doc.get());
    if (! root || ! yyjson_is_obj(root))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const auto pluginId = TryGetJsonString(root, "pluginId");
    if (! pluginId.has_value() || ! OrdinalString::EqualsNoCase(pluginId.value(), kPluginId))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    if (const auto authMode = TryGetJsonString(root, "authMode"); authMode.has_value() && ! OrdinalString::EqualsNoCase(authMode.value(), L"oauth2Pkce"))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    SecureWipe::SecureClear(outConnection.refreshToken);
    outConnection                  = {};
    outConnection.connectionName   = std::wstring(connectionName);
    outConnection.canonicalPath    = NormalizePluginPath(connectionPath);
    outConnection.pageSize         = defaults.pageSize;
    outConnection.connectTimeoutMs = defaults.connectTimeoutMs;
    outConnection.requestTimeoutMs = defaults.requestTimeoutMs;

    std::wstring explicitClientId;
    bool useDefaultClientId = true;

    if (yyjson_val* extra = yyjson_obj_get(root, "extra"); extra && yyjson_is_obj(extra))
    {
        if (const auto value = TryGetJsonString(extra, "rootKind"); value.has_value() && ! value->empty())
        {
            outConnection.rootKind = value.value();
        }
        if (const auto value = TryGetJsonString(extra, "sharedDriveId"); value.has_value())
        {
            outConnection.sharedDriveId = value.value();
        }
        if (const auto value = TryGetJsonString(extra, "googleDocsMode"); value.has_value() && ! value->empty())
        {
            outConnection.googleDocsMode = value.value();
        }
        if (const auto value = TryGetJsonBool(extra, "readOnly"); value.has_value())
        {
            outConnection.readOnly = value.value();
        }
        if (const auto value = TryGetJsonBool(extra, "useDefaultClientId"); value.has_value())
        {
            useDefaultClientId = value.value();
        }
        if (const auto value = TryGetJsonString(extra, "clientId"); value.has_value())
        {
            explicitClientId = value.value();
        }
    }

    if (useDefaultClientId)
    {
        outConnection.clientId = defaults.defaultClientId;
        if (outConnection.clientId.empty() && ! explicitClientId.empty())
        {
            outConnection.clientId = std::move(explicitClientId);
        }
    }
    else
    {
        outConnection.clientId = std::move(explicitClientId);
    }

    if (outConnection.clientId.empty())
    {
        Debug::Warning(L"GDrive: missing OAuth client id for connection '{}'. Configure plugin defaultClientId or connection extra.clientId.",
                       outConnection.connectionName);
        ShowMissingClientIdAlert();
        return HRESULT_FROM_WIN32(ERROR_BAD_CONFIGURATION);
    }

    if (! OrdinalString::EqualsNoCase(outConnection.rootKind, L"myDrive") && ! OrdinalString::EqualsNoCase(outConnection.rootKind, L"sharedDrive"))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if (OrdinalString::EqualsNoCase(outConnection.rootKind, L"sharedDrive") && outConnection.sharedDriveId.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    outConnection.connectionKey = ConnectionKeyForCache(outConnection.connectionName, outConnection.clientId);

    if (acquireSecrets)
    {
        wchar_t* rawSecret = nullptr;
        const HRESULT secretHr =
            hostConnections->GetConnectionSecret(std::wstring(connectionName).c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, nullptr, &rawSecret);
        wil::unique_cotaskmem_string refreshToken(rawSecret);
        if (FAILED(secretHr))
        {
            return secretHr;
        }
        if (! refreshToken.get() || refreshToken.get()[0] == L'\0')
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_AUTHENTICATED);
        }

        outConnection.refreshToken = refreshToken.get();
    }

    return S_OK;
}

HRESULT FileSystemGoogleDrive::GetAccessToken(const ResolvedConnection& connection, std::wstring& accessToken)
{
    SecureWipe::SecureClear(accessToken);
    if (connection.clientId.empty() || connection.refreshToken.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_AUTHENTICATED);
    }

    uint64_t cacheGeneration = 0u;
    bool waitedForRefresh    = false;
    {
        std::unique_lock lock(_tokenMutex);
        while (_tokenRefreshesInFlight.contains(connection.connectionKey))
        {
            waitedForRefresh = true;
            _tokenCv.wait(lock);
        }

        const uint64_t now = GetTickCount64();
        const auto it      = _accessTokensByConnectionKey.find(connection.connectionKey);
        if (it != _accessTokensByConnectionKey.end() && ! it->second.token.empty() && it->second.expiresAtTickMs > now + 30'000ull)
        {
            accessToken = it->second.token;
            return S_OK;
        }

        if (waitedForRefresh)
        {
            const auto lastRefresh = _lastTokenRefreshStatus.find(connection.connectionKey);
            if (lastRefresh != _lastTokenRefreshStatus.end() && FAILED(lastRefresh->second))
            {
                return lastRefresh->second;
            }
        }

        cacheGeneration = _tokenCacheGeneration;
        _tokenRefreshesInFlight.emplace(connection.connectionKey);
    }

    bool refreshCompleted     = false;
    const auto releaseRefresh = wil::scope_exit([&]() noexcept
    {
        if (refreshCompleted)
        {
            return;
        }
        {
            std::lock_guard lock(_tokenMutex);
            _lastTokenRefreshStatus[connection.connectionKey] = E_FAIL;
            _tokenRefreshesInFlight.erase(connection.connectionKey);
        }
        _tokenCv.notify_all();
    });

    AccessTokenCacheEntry refreshedEntry{};
    const HRESULT refreshHr = [&]() -> HRESULT
    {
        std::string clientIdUtf8         = Utf8FromUtf16(connection.clientId);
        std::string refreshTokenUtf8     = Utf8FromUtf16(connection.refreshToken);
        const auto clearConvertedSecrets = wil::scope_exit([&]() noexcept { SecureWipe::SecureClear(refreshTokenUtf8); });
        if (clientIdUtf8.empty() || refreshTokenUtf8.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }

        std::string body;
        body.reserve(clientIdUtf8.size() + refreshTokenUtf8.size() + 64u);
        body.append("client_id=");
        body.append(UrlEncodeUtf8(clientIdUtf8));
        body.append("&grant_type=refresh_token&refresh_token=");
        body.append(UrlEncodeUtf8(refreshTokenUtf8));
        const auto clearRequestBody = wil::scope_exit([&]() noexcept { SecureWipe::SecureClear(body); });

        const std::vector<std::string> headers = {
            "Content-Type: application/x-www-form-urlencoded",
            "Accept: application/json",
        };

        HttpResponse response{};
        const auto clearResponseBody = wil::scope_exit([&]() noexcept { SecureWipe::SecureClear(response.body); });
        HRESULT hr = PerformHttpRequest("POST", TokenEndpointUrl(), headers, body, connection.connectTimeoutMs, connection.requestTimeoutMs, response);
        if (FAILED(hr))
        {
            return hr;
        }
        if (response.statusCode < 200 || response.statusCode >= 300)
        {
            return MapHttpStatusToHresult(response.statusCode);
        }

        unique_yyjson_doc doc(yyjson_read(response.body.data(), response.body.size(), YYJSON_READ_ALLOW_BOM));
        if (! doc)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        yyjson_val* root = yyjson_doc_get_root(doc.get());
        if (! root || ! yyjson_is_obj(root))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const auto token = TryGetJsonString(root, "access_token");
        if (! token.has_value() || token->empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        uint64_t expiresIn = 3600u;
        if (const auto value = TryGetJsonUInt64Flexible(root, "expires_in"); value.has_value() && value.value() > 0u)
        {
            expiresIn = value.value();
        }
        const uint64_t validSeconds    = expiresIn > 60u ? (std::min)(expiresIn - 60u, (std::numeric_limits<uint64_t>::max)() / 1000u) : 30u;
        const uint64_t validForMs      = validSeconds * 1000u;
        refreshedEntry.token           = token.value();
        refreshedEntry.expiresAtTickMs = Common::Paging::DeadlineFromNow(GetTickCount64(), validForMs);
        return S_OK;
    }();

    if (SUCCEEDED(refreshHr))
    {
        accessToken = refreshedEntry.token;
    }

    {
        std::lock_guard lock(_tokenMutex);
        if (SUCCEEDED(refreshHr) && cacheGeneration == _tokenCacheGeneration)
        {
            _accessTokensByConnectionKey[connection.connectionKey] = std::move(refreshedEntry);
        }
        _lastTokenRefreshStatus[connection.connectionKey] = refreshHr;
        _tokenRefreshesInFlight.erase(connection.connectionKey);
    }
    _tokenCv.notify_all();
    refreshCompleted = true;

    return refreshHr;
}

HRESULT FileSystemGoogleDrive::ListChildren(const ResolvedConnection& connection, std::wstring_view parentId, std::vector<GoogleItem>& items)
{
    items.clear();
    if (parentId.empty())
    {
        return E_INVALIDARG;
    }

    const std::string parentIdUtf8 = Utf8FromUtf16(parentId);
    if (parentIdUtf8.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    std::string pageToken;
    const uint64_t pagingDurationMs = std::clamp<uint64_t>(static_cast<uint64_t>(connection.requestTimeoutMs) * 10u, 60'000u, 600'000u);
    Common::Paging::Utf8ContinuationGuard pager(Common::Paging::Limits{
        .deadlineTickMs = Common::Paging::DeadlineFromNow(GetTickCount64(), pagingDurationMs),
    });
    bool firstPage = true;
    do
    {
        const HRESULT pageBoundaryHr = firstPage ? pager.BeginFirstPage(GetTickCount64()) : pager.BeginContinuation(pageToken, GetTickCount64());
        firstPage                    = false;
        if (FAILED(pageBoundaryHr))
        {
            return pageBoundaryHr;
        }

        std::string url = FilesEndpointUrl();
        AppendQueryParam(url, "fields", std::format("nextPageToken,files({})", kDriveItemFields));
        AppendQueryParam(url, "pageSize", std::to_string(connection.pageSize));
        AppendQueryParam(url, "supportsAllDrives", "true");
        AppendQueryParam(url, "includeItemsFromAllDrives", "true");
        AppendQueryParam(url, "spaces", "drive");
        if (OrdinalString::EqualsNoCase(connection.rootKind, L"sharedDrive"))
        {
            AppendQueryParam(url, "corpora", "drive");
            AppendQueryParam(url, "driveId", Utf8FromUtf16(connection.sharedDriveId));
        }
        else
        {
            AppendQueryParam(url, "corpora", "user");
        }

        const std::string query = std::format("trashed = false and '{}' in parents", parentIdUtf8);
        AppendQueryParam(url, "q", query);
        if (! pageToken.empty())
        {
            AppendQueryParam(url, "pageToken", pageToken);
        }

        std::string body;
        HRESULT hr = PerformAuthorizedJsonGet(connection, url, body);
        if (FAILED(hr))
        {
            return hr;
        }

        unique_yyjson_doc doc(yyjson_read(body.data(), body.size(), YYJSON_READ_ALLOW_BOM));
        if (! doc)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        yyjson_val* root = yyjson_doc_get_root(doc.get());
        if (! root || ! yyjson_is_obj(root))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        pageToken.clear();
        if (const auto nextPage = TryGetJsonUtf8String(root, "nextPageToken"); nextPage.has_value())
        {
            pageToken = nextPage.value();
        }

        yyjson_val* files      = yyjson_obj_get(root, "files");
        const size_t pageItems = files && yyjson_is_arr(files) ? yyjson_arr_size(files) : 0u;
        const HRESULT pageHr   = pager.CompletePage(pageItems, body.size(), ! pageToken.empty(), pageToken, GetTickCount64());
        if (FAILED(pageHr))
        {
            return pageHr;
        }
        if (! files || ! yyjson_is_arr(files))
        {
            continue;
        }

        size_t index      = 0;
        size_t max        = 0;
        yyjson_val* entry = nullptr;
        yyjson_arr_foreach(files, index, max, entry)
        {
            if (! entry || ! yyjson_is_obj(entry))
            {
                continue;
            }

            GoogleItem item{};
            if (FAILED(ParseGoogleItem(entry, item)))
            {
                continue; // trashed or malformed entries are not exposed
            }
            items.push_back(std::move(item));
        }
    } while (! pageToken.empty());

    return S_OK;
}

HRESULT FileSystemGoogleDrive::ResolveItemByPath(const ResolvedConnection& connection, std::wstring_view canonicalPath, GoogleItem& item)
{
    item            = {};
    item.id         = OrdinalString::EqualsNoCase(connection.rootKind, L"sharedDrive") ? connection.sharedDriveId : std::wstring(L"root");
    item.name       = L"/";
    item.attributes = FILE_ATTRIBUTE_DIRECTORY;
    item.isFolder   = true;

    const std::wstring normalized = NormalizePluginPath(canonicalPath);
    if (normalized == L"/")
    {
        return S_OK;
    }

    std::wstring_view rest(normalized);
    while (! rest.empty() && rest.front() == L'/')
    {
        rest.remove_prefix(1);
    }

    while (! rest.empty())
    {
        const size_t slashPos           = rest.find(L'/');
        const std::wstring_view segment = (slashPos == std::wstring_view::npos) ? rest : rest.substr(0, slashPos);
        rest                            = (slashPos == std::wstring_view::npos) ? std::wstring_view{} : rest.substr(slashPos + 1u);

        if (segment.empty())
        {
            continue;
        }

        if (! item.isFolder)
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }

        std::vector<GoogleItem> children;
        HRESULT hr = ListChildren(connection, item.id, children);
        if (FAILED(hr))
        {
            return hr;
        }

        const GoogleItem* match = nullptr;
        hr                      = ResolveChildByExposedName(children, segment, match);
        if (FAILED(hr))
        {
            return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) ? HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) : hr;
        }

        item = *match;
    }

    return S_OK;
}

HRESULT FileSystemGoogleDrive::FetchDriveInfoPayload(const ResolvedConnection& connection, DriveInfoPayload& payload)
{
    payload = {};

    std::string aboutUrl = AboutEndpointUrl();
    AppendQueryParam(aboutUrl, "fields", "user(displayName,emailAddress),storageQuota(limit,usage,usageInDrive)");
    std::string body;
    HRESULT hr = PerformAuthorizedJsonGet(connection, aboutUrl, body);
    if (FAILED(hr))
    {
        return hr;
    }

    unique_yyjson_doc doc(yyjson_read(body.data(), body.size(), YYJSON_READ_ALLOW_BOM));
    if (! doc)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    yyjson_val* root = yyjson_doc_get_root(doc.get());
    if (! root || ! yyjson_is_obj(root))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if (yyjson_val* user = yyjson_obj_get(root, "user"); user && yyjson_is_obj(user))
    {
        payload.userDisplayName = TryGetJsonString(user, "displayName").value_or(L"");
        payload.userEmail       = TryGetJsonString(user, "emailAddress").value_or(L"");
    }

    if (! OrdinalString::EqualsNoCase(connection.rootKind, L"sharedDrive"))
    {
        payload.driveName = L"My Drive";
        if (yyjson_val* quota = yyjson_obj_get(root, "storageQuota"); quota && yyjson_is_obj(quota))
        {
            const auto limit = TryGetJsonUInt64Flexible(quota, "limit");
            const auto usage = TryGetJsonUInt64Flexible(quota, "usageInDrive").or_else([&]() { return TryGetJsonUInt64Flexible(quota, "usage"); });
            if (limit.has_value() && limit.value() > 0)
            {
                payload.totalBytes = limit.value();
                payload.hasTotal   = true;
                if (usage.has_value())
                {
                    payload.usedBytes = usage.value();
                    payload.hasUsed   = true;
                    payload.freeBytes = (usage.value() < limit.value()) ? (limit.value() - usage.value()) : 0;
                    payload.hasFree   = true;
                }
            }
        }

        return S_OK;
    }

    std::string driveUrl = DrivesEndpointUrl() + UrlEncodeUtf8(Utf8FromUtf16(connection.sharedDriveId));
    AppendQueryParam(driveUrl, "fields", "name");
    body.clear();
    hr = PerformAuthorizedJsonGet(connection, driveUrl, body);
    if (FAILED(hr))
    {
        return hr;
    }

    unique_yyjson_doc driveDoc(yyjson_read(body.data(), body.size(), YYJSON_READ_ALLOW_BOM));
    if (! driveDoc)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    yyjson_val* driveRoot = yyjson_doc_get_root(driveDoc.get());
    if (! driveRoot || ! yyjson_is_obj(driveRoot))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    payload.driveName = TryGetJsonString(driveRoot, "name").value_or(L"Shared Drive");
    return S_OK;
}

HRESULT FileSystemGoogleDrive::ReadDirectoryInfoImpl(const wchar_t* path, IFilesInformation** ppFilesInformation)
{
    if (! path || ! ppFilesInformation)
    {
        return E_POINTER;
    }

    *ppFilesInformation = nullptr;

    const std::wstring normalized = NormalizePluginPath(path);
    if (normalized == L"/")
    {
        auto* info = new (std::nothrow) FilesInformationGoogleDrive();
        if (! info)
        {
            return E_OUTOFMEMORY;
        }

        const HRESULT buildHr = info->BuildFromEntries({});
        if (FAILED(buildHr))
        {
            info->Release();
            return buildHr;
        }

        *ppFilesInformation = info;
        return S_OK;
    }

    ResolvedConnection connection{};
    HRESULT hr = ResolveConnection(path, true, connection);
    if (FAILED(hr))
    {
        return hr;
    }

    GoogleItem folder{};
    hr = ResolveItemByPath(connection, connection.canonicalPath, folder);
    if (FAILED(hr))
    {
        return hr;
    }
    if (! folder.isFolder)
    {
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
    }

    std::vector<GoogleItem> children;
    hr = ListChildren(connection, folder.id, children);
    if (FAILED(hr))
    {
        return hr;
    }

    std::map<std::wstring, size_t, std::less<>> counts;
    for (const GoogleItem& child : children)
    {
        ++counts[child.name];
    }

    std::vector<FilesInformationGoogleDrive::Entry> entries;
    entries.reserve(children.size());

    unsigned long fileIndex = 1;
    for (const GoogleItem& child : children)
    {
        FilesInformationGoogleDrive::Entry entry{};
        const auto countIt = counts.find(child.name);
        entry.name         = MakeExposedItemName(child.name, child.id, countIt == counts.end() ? 0u : countIt->second);

        entry.fileIndex      = fileIndex++;
        entry.attributes     = child.attributes;
        entry.sizeBytes      = child.sizeBytes;
        entry.creationTime   = child.creationTime;
        entry.lastAccessTime = child.lastAccessTime;
        entry.lastWriteTime  = child.lastWriteTime;
        entry.changeTime     = child.changeTime;
        entries.push_back(std::move(entry));
    }

    auto* info = new (std::nothrow) FilesInformationGoogleDrive();
    if (! info)
    {
        return E_OUTOFMEMORY;
    }

    hr = info->BuildFromEntries(std::move(entries));
    if (FAILED(hr))
    {
        info->Release();
        return hr;
    }

    *ppFilesInformation = info;
    return S_OK;
}

HRESULT FileSystemGoogleDrive::GetDriveInfoImpl(const wchar_t* path, DriveInfo* info)
{
    if (! path || ! info)
    {
        return E_POINTER;
    }

    *info = {};

    const std::wstring normalized = NormalizePluginPath(path);
    if (normalized == L"/")
    {
        std::scoped_lock lock(_stateMutex);
        _driveDisplayName = L"gdrive:/";
        _driveVolumeLabel = L"Google Drive";
        _driveFileSystem  = L"Google Drive";

        info->flags       = static_cast<DriveInfoFlags>(DRIVE_INFO_FLAG_HAS_DISPLAY_NAME | DRIVE_INFO_FLAG_HAS_VOLUME_LABEL | DRIVE_INFO_FLAG_HAS_FILE_SYSTEM);
        info->displayName = _driveDisplayName.c_str();
        info->volumeLabel = _driveVolumeLabel.c_str();
        info->fileSystem  = _driveFileSystem.c_str();
        return S_OK;
    }

    ResolvedConnection connection{};
    HRESULT hr = ResolveConnection(path, true, connection);
    if (FAILED(hr))
    {
        return hr;
    }

    GoogleItem item{};
    hr = ResolveItemByPath(connection, connection.canonicalPath, item);
    if (FAILED(hr))
    {
        return hr;
    }

    DriveInfoPayload payload{};
    hr = FetchDriveInfoPayload(connection, payload);
    if (FAILED(hr))
    {
        return hr;
    }

    std::scoped_lock lock(_stateMutex);
    _driveDisplayName = BuildDriveDisplayName(connection.connectionName, connection.canonicalPath);
    _driveVolumeLabel = payload.driveName.empty() ? std::wstring(L"Google Drive") : payload.driveName;
    _driveFileSystem  = L"Google Drive";

    uint32_t flags    = DRIVE_INFO_FLAG_HAS_DISPLAY_NAME | DRIVE_INFO_FLAG_HAS_VOLUME_LABEL | DRIVE_INFO_FLAG_HAS_FILE_SYSTEM;
    info->displayName = _driveDisplayName.c_str();
    info->volumeLabel = _driveVolumeLabel.c_str();
    info->fileSystem  = _driveFileSystem.c_str();

    if (payload.hasTotal)
    {
        info->totalBytes = payload.totalBytes;
        flags |= DRIVE_INFO_FLAG_HAS_TOTAL_BYTES;
    }
    if (payload.hasFree)
    {
        info->freeBytes = payload.freeBytes;
        flags |= DRIVE_INFO_FLAG_HAS_FREE_BYTES;
    }
    if (payload.hasUsed)
    {
        info->usedBytes = payload.usedBytes;
        flags |= DRIVE_INFO_FLAG_HAS_USED_BYTES;
    }

    info->flags = static_cast<DriveInfoFlags>(flags);
    return S_OK;
}

#if defined(_DEBUG)
namespace
{
enum class GoogleDriveDebugHttpMode
{
    Normal,
    RetryThenSuccess,
    MutationFailureThenSuccess,
    RepeatedPageToken,
    IdentityItems,
    OversizedBody,
};

struct GoogleDriveDebugHttpContext final
{
    GoogleDriveDebugHttpContext()                                              = default;
    GoogleDriveDebugHttpContext(const GoogleDriveDebugHttpContext&)            = delete;
    GoogleDriveDebugHttpContext& operator=(const GoogleDriveDebugHttpContext&) = delete;
    GoogleDriveDebugHttpContext(GoogleDriveDebugHttpContext&&)                 = delete;
    GoogleDriveDebugHttpContext& operator=(GoogleDriveDebugHttpContext&&)      = delete;

    std::mutex mutex;
    std::condition_variable cv;
    GoogleDriveDebugHttpMode mode = GoogleDriveDebugHttpMode::Normal;
    unsigned int tokenRequests    = 0u;
    unsigned int dataRequests     = 0u;
    bool tokenRequestEntered      = false;
    bool releaseTokenRequest      = true;
};

[[nodiscard]] HRESULT GoogleDriveDebugHttpHook(void* cookie,
                                               std::string_view method,
                                               std::string_view url,
                                               [[maybe_unused]] const std::vector<std::string>& headers,
                                               [[maybe_unused]] std::string_view body,
                                               HttpResponse& response) noexcept
{
    auto* context = static_cast<GoogleDriveDebugHttpContext*>(cookie);
    if (! context)
    {
        return E_POINTER;
    }

    if (method == "POST" && url == kTokenEndpoint)
    {
        std::unique_lock lock(context->mutex);
        ++context->tokenRequests;
        context->tokenRequestEntered = true;
        context->cv.notify_all();
        context->cv.wait(lock, [&]() noexcept { return context->releaseTokenRequest; });
        response.statusCode = 200;
        response.body       = R"json({"access_token":"debug-access-token","expires_in":3600})json";
        return S_OK;
    }

    unsigned int requestNumber = 0u;
    GoogleDriveDebugHttpMode mode{};
    {
        std::lock_guard lock(context->mutex);
        requestNumber = ++context->dataRequests;
        mode          = context->mode;
    }

    switch (mode)
    {
        case GoogleDriveDebugHttpMode::RetryThenSuccess:
            if (requestNumber <= 2u)
            {
                response.statusCode = 429;
                response.retryAfter = "0";
                response.body       = R"json({"error":"rateLimit"})json";
            }
            else
            {
                response.statusCode = 200;
                response.body       = R"json({"ok":true})json";
            }
            return S_OK;
        case GoogleDriveDebugHttpMode::MutationFailureThenSuccess:
            if (requestNumber == 1u)
            {
                response.statusCode = 503;
                response.body       = R"json({"error":"commit state unknown"})json";
            }
            else
            {
                response.statusCode = 200;
                response.body       = R"json({"id":"duplicate","name":"folder","mimeType":"application/vnd.google-apps.folder","version":"1"})json";
            }
            return S_OK;
        case GoogleDriveDebugHttpMode::RepeatedPageToken:
            response.statusCode = 200;
            response.body       = R"json({"nextPageToken":"repeat","files":[]})json";
            return S_OK;
        case GoogleDriveDebugHttpMode::IdentityItems:
            response.statusCode = 200;
            response.body =
                R"json({"files":[{"id":"AbC","name":"same","mimeType":"application/octet-stream"},{"id":"abc","name":"same","mimeType":"application/octet-stream"},{"id":"literal","name":"same [id:AbC]","mimeType":"application/octet-stream"},{"id":"upper","name":"Case","mimeType":"application/octet-stream"},{"id":"lower","name":"case","mimeType":"application/octet-stream"}]})json";
            return S_OK;
        case GoogleDriveDebugHttpMode::OversizedBody:
            response.statusCode = 200;
            response.body.assign(kMaxJsonResponseBytes + 1u, 'x');
            return S_OK;
        case GoogleDriveDebugHttpMode::Normal:
        default:
            response.statusCode = 200;
            response.body       = R"json({"files":[]})json";
            return S_OK;
    }
}
} // namespace

HRESULT FileSystemGoogleDrive::RunDebugSelfTests(unsigned int* passed, unsigned int* failed) noexcept
{
    if (! passed || ! failed)
    {
        return E_POINTER;
    }
    *passed = 0u;
    *failed = 0u;
    constexpr Common::DebugSelfTest::Check check{L"Google Drive"};

    try
    {
        check(MakeExposedItemName(L"same", L"AbC", 2u) == L"same [id:AbC]" && MakeExposedItemName(L"same", L"abc", 2u) == L"same [id:abc]",
              L"case-distinct opaque IDs remain distinct in duplicate-name display paths",
              *passed,
              *failed);
        check(MakeExposedItemName(L"same [id:AbC]", L"literal", 1u) == L"same [id:AbC] [id:literal]",
              L"literal suffix-like names receive a reversible extra identity decoration",
              *passed,
              *failed);

        HttpResponse tinyResponse{};
        CurlResponseWriteState tinyWrite{.response = &tinyResponse, .maxBytes = 4u};
        char fiveBytes[] = "12345";
        check(CurlWriteToString(fiveBytes, 1u, 5u, &tinyWrite) == 0u && tinyWrite.tooLarge && tinyResponse.body.empty(),
              L"response callback rejects a body beyond its hard cap before retaining bytes",
              *passed,
              *failed);

        CurlTransferDeadline expired{.deadlineTickMs = GetTickCount64()};
        check(CurlCheckTransferDeadline(&expired, 0, 0, 0, 0) == 1 && expired.abortStatus == HRESULT_FROM_WIN32(ERROR_TIMEOUT),
              L"transfer progress callback aborts trickle traffic at the hard deadline",
              *passed,
              *failed);

        FileSystemGoogleDrive fs(nullptr);
        ResolvedConnection connection{};
        connection.connectionName   = L"debug";
        connection.connectionKey    = L"debug|client";
        connection.clientId         = L"client";
        connection.refreshToken     = L"refresh";
        connection.requestTimeoutMs = 5'000u;
        connection.pageSize         = 200u;

        GoogleDriveDebugHttpContext context{};
        context.releaseTokenRequest = false;
        DebugHttpRequestHookScope hook(GoogleDriveDebugHttpHook, &context);
        constexpr size_t workerCount = 8u;
        std::barrier start(static_cast<std::ptrdiff_t>(workerCount + 1u));
        std::array<HRESULT, workerCount> results{};
        std::array<std::wstring, workerCount> tokens{};
        std::vector<std::jthread> workers;
        workers.reserve(workerCount);
        for (size_t index = 0u; index < workerCount; ++index)
        {
            workers.emplace_back([&, index]()
            {
                start.arrive_and_wait();
                results[index] = fs.GetAccessToken(connection, tokens[index]);
            });
        }
        start.arrive_and_wait();
        {
            std::unique_lock lock(context.mutex);
            context.cv.wait(lock, [&]() noexcept { return context.tokenRequestEntered; });
            context.releaseTokenRequest = true;
        }
        context.cv.notify_all();
        for (std::jthread& worker : workers)
        {
            worker.join();
        }
        const bool allSharedOneToken = std::ranges::all_of(results, [](HRESULT hr) noexcept { return hr == S_OK; }) &&
                                       std::ranges::all_of(tokens, [](const std::wstring& token) noexcept { return token == L"debug-access-token"; });
        check(allSharedOneToken && context.tokenRequests == 1u, L"concurrent callers share one access-token refresh request", *passed, *failed);

        {
            std::lock_guard lock(context.mutex);
            context.mode         = GoogleDriveDebugHttpMode::RetryThenSuccess;
            context.dataRequests = 0u;
        }
        g_debugSuppressRetrySleep.store(true, std::memory_order_release);
        std::string responseBody;
        HRESULT hr = fs.PerformAuthorizedJsonGet(connection, "https://www.googleapis.com/drive/v3/files", responseBody);
        g_debugSuppressRetrySleep.store(false, std::memory_order_release);
        check(hr == S_OK && context.dataRequests == 3u && responseBody == R"json({"ok":true})json",
              L"authorized GET applies bounded 429 retry and then returns the successful body",
              *passed,
              *failed);

        {
            std::lock_guard lock(context.mutex);
            context.mode         = GoogleDriveDebugHttpMode::MutationFailureThenSuccess;
            context.dataRequests = 0u;
        }
        GoogleItem mutationResult;
        hr = fs.CreateFolderItem(connection, L"root", L"folder", mutationResult);
        check(FAILED(hr) && context.dataRequests == 1u && mutationResult.id.empty(),
              L"an authorized mutation that returns 5xx is attempted once because its commit state is unknown",
              *passed,
              *failed);

        {
            std::lock_guard lock(context.mutex);
            context.mode         = GoogleDriveDebugHttpMode::OversizedBody;
            context.dataRequests = 0u;
        }
        hr = fs.PerformAuthorizedJsonGet(connection, "https://www.googleapis.com/drive/v3/files", responseBody);
        check(hr == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE) && responseBody.empty(),
              L"authorized GET rejects an oversized JSON response body",
              *passed,
              *failed);

        {
            std::lock_guard lock(context.mutex);
            context.mode         = GoogleDriveDebugHttpMode::RepeatedPageToken;
            context.dataRequests = 0u;
        }
        std::vector<GoogleItem> items;
        hr = fs.ListChildren(connection, L"root", items);
        check(hr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA) && context.dataRequests == 2u,
              L"Google paging rejects a repeated token before a third request",
              *passed,
              *failed);

        {
            std::lock_guard lock(context.mutex);
            context.mode         = GoogleDriveDebugHttpMode::IdentityItems;
            context.dataRequests = 0u;
        }
        GoogleItem item{};
        const HRESULT upperIdHr   = fs.ResolveItemByPath(connection, L"/same [id:AbC]", item);
        const bool upperId        = upperIdHr == S_OK && item.id == L"AbC";
        const HRESULT lowerIdHr   = fs.ResolveItemByPath(connection, L"/same [id:abc]", item);
        const bool lowerId        = lowerIdHr == S_OK && item.id == L"abc";
        const HRESULT literalHr   = fs.ResolveItemByPath(connection, L"/same [id:AbC] [id:literal]", item);
        const bool literal        = literalHr == S_OK && item.id == L"literal";
        const HRESULT upperNameHr = fs.ResolveItemByPath(connection, L"/Case", item);
        const bool upperName      = upperNameHr == S_OK && item.id == L"upper";
        const HRESULT lowerNameHr = fs.ResolveItemByPath(connection, L"/case", item);
        const bool lowerName      = lowerNameHr == S_OK && item.id == L"lower";
        check(upperId && lowerId && literal && upperName && lowerName,
              L"enumeration identity round-trips literal decorations, case-distinct IDs, and case-distinct names",
              *passed,
              *failed);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception& error)
    {
        // Exported debug-test ABI boundary: convert named standard failures into a deterministic red result.
        Debug::Error(L"Google Drive debug selftests failed with std::exception: {}", Utf16FromUtf8(error.what()));
        ++*failed;
    }

    g_debugSuppressRetrySleep.store(false, std::memory_order_release);
    return *failed == 0u ? S_OK : E_FAIL;
}

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderGoogleDriveDebugSelfTests(unsigned int* passed, unsigned int* failed)
{
    const HRESULT hr = FileSystemGoogleDrive::RunDebugSelfTests(passed, failed);
#if defined(ENABLE_TESTS)
    unsigned int containmentPassed = 0u;
    unsigned int containmentFailed = 0u;
    FileSystemGoogleDriveSelfTest::RunDriveStalledRequestCancelSelfTests(containmentPassed, containmentFailed);
    FileSystemGoogleDriveSelfTest::RunDriveZeroTimestampReplaceSelfTests(containmentPassed, containmentFailed);
    if (passed != nullptr && failed != nullptr)
    {
        *passed += containmentPassed;
        *failed += containmentFailed;
    }
    return FAILED(hr) || containmentFailed != 0u ? E_FAIL : S_OK;
#else
    return hr;
#endif
}
#endif

// ---------------------------------------------------------------------------------------------
// R0f-GDrive: authorized requests, streaming I/O, directory operations, and the test seam.
// ---------------------------------------------------------------------------------------------
HRESULT FileSystemGoogleDrive::PerformAuthorizedRequest(const ResolvedConnection& connection,
                                                        std::string_view method,
                                                        std::string_view url,
                                                        const std::vector<std::string>& extraHeaders,
                                                        std::string_view body,
                                                        size_t maxResponseBytes,
                                                        long& statusCodeOut,
                                                        std::string& bodyOut,
                                                        std::string* locationOut,
                                                        std::string* rangeOut,
                                                        AuthorizedRequestRetry retryPolicy)
{
    statusCodeOut = 0;
    bodyOut.clear();
    if (locationOut != nullptr)
    {
        locationOut->clear();
    }
    if (rangeOut != nullptr)
    {
        rangeOut->clear();
    }

    const uint64_t operationDurationMs = std::clamp<uint64_t>(static_cast<uint64_t>(connection.requestTimeoutMs) * 2u, 30'000u, 600'000u);
    const uint64_t deadlineTickMs      = Common::Paging::DeadlineFromNow(GetTickCount64(), operationDurationMs);
    unsigned int retryCount            = 0u;
    bool refreshedAfterUnauthorized    = false;

    while (true)
    {
        if (const HRESULT control = FileSystemCheckOperationControl(t_driveOperationOptions); FAILED(control))
        {
            return control;
        }
        if (GetTickCount64() >= deadlineTickMs)
        {
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }

        std::wstring accessToken;
        HRESULT hr = GetAccessToken(connection, accessToken);
        if (FAILED(hr))
        {
            return hr;
        }

        std::string accessTokenUtf8 = Utf8FromUtf16(accessToken);
        const auto clearAccessToken = wil::scope_exit([&]() noexcept
        {
            SecureWipe::SecureClear(accessToken);
            SecureWipe::SecureClear(accessTokenUtf8);
        });
        if (accessTokenUtf8.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }

        std::vector<std::string> headers;
        headers.reserve(extraHeaders.size() + 2u);
        headers.push_back(std::format("Authorization: Bearer {}", accessTokenUtf8));
        headers.emplace_back("Accept: application/json");
        headers.insert(headers.end(), extraHeaders.begin(), extraHeaders.end());
        const auto clearAuthorizationHeader = wil::scope_exit([&]() noexcept
        {
            for (std::string& header : headers)
            {
                SecureWipe::SecureClear(header);
            }
        });

        HttpResponse response{};
        const uint64_t requestStartTickMs = GetTickCount64();
        if (requestStartTickMs >= deadlineTickMs)
        {
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }
        // Each attempt is bounded by the hard request timeout (the provider-owned bound); the
        // operation budget above only caps the retry sequence as a whole.
        uint64_t attemptMs = deadlineTickMs - requestStartTickMs;
        if (connection.requestTimeoutMs != 0u)
        {
            attemptMs = (std::min<uint64_t>)(attemptMs, connection.requestTimeoutMs);
        }
        const uint32_t requestTimeoutMs = static_cast<uint32_t>((std::min<uint64_t>)(attemptMs, (std::numeric_limits<uint32_t>::max)()));
        hr = PerformHttpRequest(method, url, headers, body, connection.connectTimeoutMs, requestTimeoutMs, response, maxResponseBytes);
        if (FAILED(hr))
        {
            return hr;
        }

        if (response.statusCode == 401 && ! refreshedAfterUnauthorized)
        {
            {
                std::scoped_lock lock(_tokenMutex);
                const auto cached = _accessTokensByConnectionKey.find(connection.connectionKey);
                if (cached != _accessTokensByConnectionKey.end() && cached->second.token == accessToken)
                {
                    cached->second.expiresAtTickMs = 0u;
                }
            }
            refreshedAfterUnauthorized = true;
            continue;
        }

        // A mutation that reached Drive may have committed even when the response is 429/5xx. Only
        // read-only requests may repeat automatically; mutation callers surface the status and let
        // their identity-aware operation logic reconcile or report the uncertain outcome.
        if (retryPolicy == AuthorizedRequestRetry::ReadOnly && IsRetryableAuthorizedStatus(response.statusCode) && retryCount < kMaxAuthorizedRetries)
        {
            const uint64_t delayMs = ComputeRetryDelayMs(response, retryCount++);
            if (GetTickCount64() >= deadlineTickMs || delayMs >= deadlineTickMs - GetTickCount64())
            {
                return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
            }
#if defined(_DEBUG)
            if (! g_debugSuppressRetrySleep.load(std::memory_order_acquire))
#endif
            {
                SleepWithOperationControl(delayMs);
            }
            continue;
        }

        statusCodeOut = response.statusCode;
        bodyOut       = std::move(response.body);
        if (locationOut != nullptr)
        {
            *locationOut = std::move(response.location);
        }
        if (rangeOut != nullptr)
        {
            *rangeOut = std::move(response.range);
        }
        return S_OK;
    }
}

HRESULT FileSystemGoogleDrive::PerformAuthorizedJsonGet(const ResolvedConnection& connection, std::string_view url, std::string& body)
{
    long statusCode = 0;
    const HRESULT hr =
        PerformAuthorizedRequest(connection, "GET", url, {}, {}, kMaxJsonResponseBytes, statusCode, body, nullptr, nullptr, AuthorizedRequestRetry::ReadOnly);
    if (FAILED(hr))
    {
        body.clear();
        return hr;
    }
    if (const HRESULT statusHr = MapHttpStatusToHresult(statusCode); FAILED(statusHr))
    {
        body.clear();
        return statusHr;
    }
    return S_OK;
}

HRESULT FileSystemGoogleDrive::DownloadRange(const ResolvedConnection& connection, std::wstring_view id, uint64_t offset, size_t bytes, std::string& data)
{
    data.clear();
    if (id.empty() || bytes == 0u)
    {
        return E_INVALIDARG;
    }
    std::string url = FilesEndpointUrl() + "/" + UrlEncodeUtf8(Utf8FromUtf16(id));
    AppendQueryParam(url, "alt", "media");
    AppendQueryParam(url, "supportsAllDrives", "true");
    const std::vector<std::string> headers{std::format("Range: bytes={}-{}", offset, offset + bytes - 1u)};
    long statusCode = 0;
    std::string body;
    const HRESULT hr =
        PerformAuthorizedRequest(connection, "GET", url, headers, {}, bytes + 1u, statusCode, body, nullptr, nullptr, AuthorizedRequestRetry::ReadOnly);
    if (FAILED(hr))
    {
        return hr == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE) ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA) : hr;
    }
    if (statusCode != 206 && statusCode != 200)
    {
        const HRESULT statusHr = MapHttpStatusToHresult(statusCode);
        if (FAILED(statusHr))
        {
            return statusHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ? HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) : statusHr;
        }
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    // A 200 answers the whole object: only acceptable when the range started at the beginning.
    if ((statusCode == 200 && offset != 0u) || body.size() != bytes)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    data = std::move(body);
    return S_OK;
}

HRESULT FileSystemGoogleDrive::UploadResumable(const ResolvedConnection& connection,
                                               std::wstring_view existingId,
                                               std::wstring_view parentId,
                                               std::wstring_view name,
                                               HANDLE file,
                                               uint64_t sizeBytes,
                                               GoogleItem& result)
{
    result = {};
    if (file == nullptr || file == INVALID_HANDLE_VALUE || (existingId.empty() && (parentId.empty() || name.empty())))
    {
        return E_INVALIDARG;
    }

    // 1) Open the session: a new object under `parentId`, or new content for `existingId`.
    std::string url    = UploadFilesEndpointUrl();
    std::string method = "POST";
    std::string metadata;
    if (existingId.empty())
    {
        metadata = std::format(R"({{"name":{},"parents":[{}]}})", JsonQuote(name), JsonQuote(parentId));
    }
    else
    {
        url += "/" + UrlEncodeUtf8(Utf8FromUtf16(existingId));
        method   = "PATCH";
        metadata = "{}";
    }
    AppendQueryParam(url, "uploadType", "resumable");
    AppendQueryParam(url, "supportsAllDrives", "true");
    AppendQueryParam(url, "fields", std::string(kDriveItemFields));
    const std::vector<std::string> sessionHeaders{"Content-Type: application/json; charset=UTF-8", std::format("X-Upload-Content-Length: {}", sizeBytes)};
    long statusCode = 0;
    std::string body;
    std::string sessionUrl;
    HRESULT hr = PerformAuthorizedRequest(
        connection, method, url, sessionHeaders, metadata, kMaxJsonResponseBytes, statusCode, body, &sessionUrl, nullptr, AuthorizedRequestRetry::Mutation);
    if (FAILED(hr))
    {
        return hr;
    }
    if (const HRESULT statusHr = MapHttpStatusToHresult(statusCode); FAILED(statusHr))
    {
        return statusHr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND) ? HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) : statusHr;
    }
    if (sessionUrl.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    // 2) Send the bytes in order. Drive acknowledges each intermediate chunk with 308 and the
    //    acknowledged range; the final chunk answers with the published resource.
    LARGE_INTEGER start{};
    if (! SetFilePointerEx(file, start, nullptr, FILE_BEGIN))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    std::string chunk;
    chunk.resize(static_cast<size_t>((std::min<uint64_t>)(sizeBytes, kDriveUploadChunkBytes)));
    uint64_t offset = 0u;
    do
    {
        if (const HRESULT control = FileSystemCheckOperationControl(t_driveOperationOptions); FAILED(control))
        {
            return control;
        }
        const size_t want = static_cast<size_t>((std::min<uint64_t>)(sizeBytes - offset, kDriveUploadChunkBytes));
        size_t have       = 0u;
        while (have < want)
        {
            DWORD read = 0u;
            if (! ReadFile(file, chunk.data() + have, static_cast<DWORD>(want - have), &read, nullptr))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            if (read == 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
            }
            have += read;
        }
        const std::string contentRange =
            sizeBytes == 0u ? std::string("Content-Range: bytes */0") : std::format("Content-Range: bytes {}-{}/{}", offset, offset + want - 1u, sizeBytes);
        const std::vector<std::string> chunkHeaders{"Content-Type: application/octet-stream", contentRange};
        std::string range;
        hr = PerformAuthorizedRequest(connection,
                                      "PUT",
                                      sessionUrl,
                                      chunkHeaders,
                                      std::string_view(chunk.data(), want),
                                      kMaxJsonResponseBytes,
                                      statusCode,
                                      body,
                                      nullptr,
                                      &range,
                                      AuthorizedRequestRetry::Mutation);
        if (FAILED(hr))
        {
            return hr;
        }
        offset += want;
        if (statusCode == 308)
        {
            // The acknowledged range must end exactly where this chunk ended.
            if (offset >= sizeBytes || range != std::format("bytes=0-{}", offset - 1u))
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            continue;
        }
        if (const HRESULT statusHr = MapHttpStatusToHresult(statusCode); FAILED(statusHr))
        {
            return statusHr;
        }
        if (offset < sizeBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA); // published before every byte was sent
        }
        return ParseGoogleItemDocument(body, result);
    } while (offset < sizeBytes);
    return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

namespace
{
constexpr uint64_t kDriveReadChunkBytes = 4ull * 1024u * 1024u;

[[nodiscard]] std::string Rfc3339FromFileTime64(__int64 ticks)
{
    FILETIME fileTime{};
    fileTime.dwLowDateTime  = static_cast<DWORD>(static_cast<unsigned __int64>(ticks) & 0xFFFFFFFFull);
    fileTime.dwHighDateTime = static_cast<DWORD>(static_cast<unsigned __int64>(ticks) >> 32u);
    SYSTEMTIME systemTime{};
    if (! FileTimeToSystemTime(&fileTime, &systemTime))
    {
        return {};
    }
    return std::format("{:04}-{:02}-{:02}T{:02}:{:02}:{:02}.{:03}Z",
                       systemTime.wYear,
                       systemTime.wMonth,
                       systemTime.wDay,
                       systemTime.wHour,
                       systemTime.wMinute,
                       systemTime.wSecond,
                       systemTime.wMilliseconds);
}

// Reads a Drive file through ranged `alt=media` requests. Every chunk fetch re-reads the object's
// revision first: a file replaced mid-read fails with ERROR_REVISION_MISMATCH instead of splicing
// two revisions into one stream.
class GoogleDriveRangedFileReader final : public IFileReader
{
public:
    GoogleDriveRangedFileReader(FileSystemGoogleDrive* owner,
                                FileSystemGoogleDrive::ResolvedConnection&& connection,
                                FileSystemGoogleDrive::GoogleItem item) noexcept
        : _owner(owner),
          _connection(std::move(connection)),
          _item(std::move(item)),
          _size(_item.sizeBytes)
    {
    }
    GoogleDriveRangedFileReader(const GoogleDriveRangedFileReader&)            = delete;
    GoogleDriveRangedFileReader(GoogleDriveRangedFileReader&&)                 = delete;
    GoogleDriveRangedFileReader& operator=(const GoogleDriveRangedFileReader&) = delete;
    GoogleDriveRangedFileReader& operator=(GoogleDriveRangedFileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }
        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileReader))
        {
            *ppvObject = static_cast<IFileReader*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }
    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG remaining = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (remaining == 0u)
        {
            delete this;
        }
        return remaining;
    }

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (! sizeBytes)
        {
            return E_POINTER;
        }
        *sizeBytes = _size;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        uint64_t base = 0u;
        switch (origin)
        {
            case FILE_BEGIN: base = 0u; break;
            case FILE_CURRENT: base = _position; break;
            case FILE_END: base = _size; break;
            default: return E_INVALIDARG;
        }
        if (offset < 0)
        {
            const uint64_t back = static_cast<uint64_t>(-(offset + 1)) + 1u;
            if (back > base)
            {
                return E_INVALIDARG;
            }
            _position = base - back;
        }
        else
        {
            if (static_cast<uint64_t>(offset) > (std::numeric_limits<uint64_t>::max)() - base)
            {
                return E_INVALIDARG;
            }
            _position = base + static_cast<uint64_t>(offset);
        }
        if (newPosition)
        {
            *newPosition = _position;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (! bytesRead)
        {
            return E_POINTER;
        }
        *bytesRead = 0u;
        if (! buffer && bytesToRead != 0u)
        {
            return E_POINTER;
        }
        if (bytesToRead == 0u || _position >= _size)
        {
            return S_OK;
        }
        try
        {
            if (_cache.empty() || _position < _cacheOffset || _position >= _cacheOffset + _cache.size())
            {
                FileSystemGoogleDrive::GoogleItem current;
                HRESULT hr = _owner->GetItemById(_connection, _item.id, current);
                if (FAILED(hr))
                {
                    return hr;
                }
                if (current.version != _item.version || current.sizeBytes != _size)
                {
                    return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
                }
                const size_t want = static_cast<size_t>((std::min<uint64_t>)(_size - _position, kDriveReadChunkBytes));
                hr                = _owner->DownloadRange(_connection, _item.id, _position, want, _cache);
                if (FAILED(hr))
                {
                    _cache.clear();
                    return hr;
                }
                _cacheOffset = _position;
            }
            const size_t within    = static_cast<size_t>(_position - _cacheOffset);
            const size_t available = _cache.size() - within;
            const size_t copy      = (std::min<size_t>)(available, bytesToRead);
            std::memcpy(buffer, _cache.data() + within, copy);
            _position += copy;
            *bytesRead = static_cast<unsigned long>(copy);
            return S_OK;
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
        catch (const std::exception&)
        {
            Debug::Error(L"Google Drive ranged reader failed with std::exception.");
            return E_FAIL;
        }
    }

private:
    ~GoogleDriveRangedFileReader() = default;

    std::atomic<ULONG> _refCount{1u};
    wil::com_ptr_nothrow<FileSystemGoogleDrive> _owner;
    FileSystemGoogleDrive::ResolvedConnection _connection;
    FileSystemGoogleDrive::GoogleItem _item;
    uint64_t _size        = 0u;
    uint64_t _position    = 0u;
    uint64_t _cacheOffset = 0u;
    std::string _cache;
};

constexpr Common::Files::DeleteOnCloseTemporaryFileOptions kDriveTemporaryFileOptions{
    .prefix             = L"rsg",
    .flagsAndAttributes = FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_SEQUENTIAL_SCAN,
};

// Stages the bytes in a delete-on-close temporary file and publishes them with one resumable
// upload at Commit: the destination name appears only when Drive acknowledges the last chunk, and
// an overwrite replaces the existing object's content under its own id.
class GoogleDriveFileWriter final : public IFileWriter,
                                    public IFileWriterExpectedSize,
                                    public IFileWriterCommitSizeProof,
                                    public IFileWriterExpectedReplacement,
                                    public IFileWriterContentProof
{
public:
    GoogleDriveFileWriter(FileSystemGoogleDrive* owner,
                          FileSystemGoogleDrive::ResolvedConnection&& connection,
                          std::wstring parentId,
                          std::wstring leafName,
                          bool allowOverwrite,
                          wil::unique_hfile tempFile,
                          FileSystemGoogleDrive::GoogleItem creationOccupant) noexcept
        : _owner(owner),
          _connection(std::move(connection)),
          _parentId(std::move(parentId)),
          _leafName(std::move(leafName)),
          _allowOverwrite(allowOverwrite),
          _tempFile(std::move(tempFile)),
          _creationOccupant(std::move(creationOccupant))
    {
    }
    GoogleDriveFileWriter(const GoogleDriveFileWriter&)            = delete;
    GoogleDriveFileWriter(GoogleDriveFileWriter&&)                 = delete;
    GoogleDriveFileWriter& operator=(const GoogleDriveFileWriter&) = delete;
    GoogleDriveFileWriter& operator=(GoogleDriveFileWriter&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }
        *ppvObject = nullptr;
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileWriter))
        {
            *ppvObject = static_cast<IFileWriter*>(this);
        }
        else if (riid == __uuidof(IFileWriterExpectedSize))
        {
            *ppvObject = static_cast<IFileWriterExpectedSize*>(this);
        }
        else if (riid == __uuidof(IFileWriterCommitSizeProof))
        {
            *ppvObject = static_cast<IFileWriterCommitSizeProof*>(this);
        }
        else if (riid == __uuidof(IFileWriterExpectedReplacement))
        {
            *ppvObject = static_cast<IFileWriterExpectedReplacement*>(this);
        }
        else if (riid == __uuidof(IFileWriterContentProof))
        {
            *ppvObject = static_cast<IFileWriterContentProof*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1u, std::memory_order_relaxed) + 1u;
    }
    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG remaining = _refCount.fetch_sub(1u, std::memory_order_acq_rel) - 1u;
        if (remaining == 0u)
        {
            delete this;
        }
        return remaining;
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
        if (bytesWritten)
        {
            *bytesWritten = 0u;
        }
        if (! buffer && bytesToWrite != 0u)
        {
            return E_POINTER;
        }
        if (_committed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        const auto* bytes  = static_cast<const std::byte*>(buffer);
        unsigned long done = 0u;
        while (done < bytesToWrite)
        {
            DWORD written = 0u;
            if (! WriteFile(_tempFile.get(), bytes + done, bytesToWrite - done, &written, nullptr))
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
            if (written == 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
            }
            done += written;
        }
        _position += done;
        if (bytesWritten)
        {
            *bytesWritten = done;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetExpectedSize(uint64_t sizeBytes) noexcept override
    {
        if (_committed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        _expectedSize    = sizeBytes;
        _hasExpectedSize = true;
        return S_OK;
    }

    // R3-1: the occupant the host showed the user. It must be the object present when the writer
    // opened and unchanged since; Commit then replaces exactly that object id at that version.
    HRESULT STDMETHODCALLTYPE SetExpectedReplacement(const FileSystemBasicInformation* expected) noexcept override
    {
        if (! expected || expected->sizeBytes < sizeof(FileSystemBasicInformation))
        {
            return E_INVALIDARG;
        }
        if (! _allowOverwrite || _committed || _position != 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        if (_creationOccupant.id.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
        // R0-RC4: a zero timestamp on either side leaves no identity to compare; refuse rather than
        // replace whichever object the writer observed when it opened.
        if (expected->lastWriteTime == 0 || _creationOccupant.lastWriteTime == 0 || expected->lastWriteTime != _creationOccupant.lastWriteTime)
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
        _replaceConditional = true;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Commit() noexcept override
    {
        if (_committed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        if (_hasExpectedSize && _expectedSize != _position)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (! FlushFileBuffers(_tempFile.get()))
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        try
        {
            // The name is decided at publication: an object that appeared since the writer opened is
            // a conflict unless overwrite was granted, in which case its content is replaced in place.
            FileSystemGoogleDrive::GoogleItem existing;
            const HRESULT existingHr = _owner->ResolveChildByName(_connection, _parentId, _leafName, existing);
            if (FAILED(existingHr) && ! IsDriveNotFound(existingHr))
            {
                return existingHr;
            }
            std::wstring existingId;
            if (SUCCEEDED(existingHr))
            {
                if (existing.isFolder || ! _allowOverwrite)
                {
                    return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                }
                // Replace only the object that was there when the writer opened, at the revision it had.
                if (_creationOccupant.id.empty())
                {
                    return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
                }
                if (existing.id != _creationOccupant.id || existing.version != _creationOccupant.version)
                {
                    return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
                }
                existingId = existing.id;
            }
            else if (_replaceConditional || ! _creationOccupant.id.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH); // the occupant vanished after the decision
            }
            FileSystemGoogleDrive::GoogleItem published;
            const HRESULT hr = _owner->UploadResumable(_connection, existingId, _parentId, _leafName, _tempFile.get(), _position, published);
            if (FAILED(hr))
            {
                return hr;
            }
            if (published.sizeBytes != _position)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            _committedSize      = _position;
            _publishedSha256Hex = published.sha256Hex; // R3-2
            _committed          = true;
            return S_OK;
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
        catch (const std::exception&)
        {
            Debug::Error(L"Google Drive writer Commit failed with std::exception.");
            return E_FAIL;
        }
    }

    // R3-2: Drive computes sha256Checksum for the binary content it stores; the published resource
    // returns it and the host compares it with the bytes it streamed.
    HRESULT STDMETHODCALLTYPE GetContentProofAlgorithms(uint32_t* algorithmMask) noexcept override
    {
        if (algorithmMask == nullptr)
        {
            return E_POINTER;
        }
        *algorithmMask = 1u << FILESYSTEM_CONTENT_PROOF_SHA256_256;
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
        if (! Common::Crypto::DecodeHexDigest(_publishedSha256Hex, digest) || digest.size() != 32u)
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        *proof                  = {};
        proof->sizeBytes        = sizeof(*proof);
        proof->algorithm        = FILESYSTEM_CONTENT_PROOF_SHA256_256;
        proof->contentSizeBytes = _committedSize;
        std::memcpy(proof->digest, digest.data(), digest.size());
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetCommittedSize(uint64_t* sizeBytes) noexcept override
    {
        if (! sizeBytes)
        {
            return E_POINTER;
        }
        if (! _committed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }
        *sizeBytes = _committedSize;
        return S_OK;
    }

private:
    ~GoogleDriveFileWriter() = default;

    std::atomic<ULONG> _refCount{1u};
    wil::com_ptr_nothrow<FileSystemGoogleDrive> _owner;
    FileSystemGoogleDrive::ResolvedConnection _connection;
    std::wstring _parentId;
    std::wstring _leafName;
    bool _allowOverwrite = false;
    wil::unique_hfile _tempFile;
    FileSystemGoogleDrive::GoogleItem _creationOccupant; // R3-1: the object under the name when the writer opened (empty id = none)
    bool _replaceConditional = false;                    // R3-1: the host granted a replacement of that occupant
    uint64_t _position       = 0u;
    uint64_t _expectedSize   = 0u;
    uint64_t _committedSize  = 0u;
    bool _hasExpectedSize    = false;
    bool _committed          = false;
    std::string _publishedSha256Hex; // R3-2: Drive's digest of the published object
};

[[nodiscard]] bool IsGoogleNativeDocument(const FileSystemGoogleDrive::GoogleItem& item) noexcept
{
    return ! item.isFolder && ! item.isShortcut && item.mimeType.starts_with(L"application/vnd.google-apps.");
}
} // namespace

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept
{
    if (! fileAttributes)
    {
        return E_POINTER;
    }
    *fileAttributes = 0u;
    if (! path)
    {
        return E_POINTER;
    }
    try
    {
        ResolvedConnection connection;
        HRESULT hr = ResolveConnection(path, true, connection);
        if (FAILED(hr))
        {
            return hr;
        }
        GoogleItem item;
        hr = ResolveItemByPath(connection, connection.canonicalPath, item);
        if (FAILED(hr))
        {
            return hr;
        }
        *fileAttributes = item.attributes;
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive GetAttributes failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept
{
    if (! reader)
    {
        return E_POINTER;
    }
    *reader = nullptr;
    if (! path)
    {
        return E_POINTER;
    }
    try
    {
        ResolvedConnection connection;
        HRESULT hr = ResolveConnection(path, true, connection);
        if (FAILED(hr))
        {
            return hr;
        }
        GoogleItem item;
        hr = ResolveItemByPath(connection, connection.canonicalPath, item);
        if (FAILED(hr))
        {
            return hr;
        }
        if (item.isFolder)
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY_NOT_SUPPORTED);
        }
        if (IsGoogleNativeDocument(item))
        {
            // Docs/Sheets/Slides have no byte stream of their own (export formats are a separate feature).
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }
        auto* instance = new (std::nothrow) GoogleDriveRangedFileReader(this, std::move(connection), std::move(item));
        if (! instance)
        {
            return E_OUTOFMEMORY;
        }
        *reader = instance;
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive CreateFileReader failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept
{
    if (! writer)
    {
        return E_POINTER;
    }
    *writer = nullptr;
    if (! path)
    {
        return E_POINTER;
    }
    try
    {
        ResolvedConnection connection;
        HRESULT hr = ResolveConnection(path, true, connection);
        if (FAILED(hr))
        {
            return hr;
        }
        if (connection.readOnly)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        GoogleItem parent;
        std::wstring leafName;
        hr = ResolveParentAndLeaf(connection, connection.canonicalPath, parent, leafName);
        if (FAILED(hr))
        {
            return hr;
        }
        if (! IsValidDriveLeafName(leafName))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
        const bool allowOverwrite = (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
        GoogleItem existing;
        const HRESULT existingHr = ResolveChildByName(connection, parent.id, leafName, existing);
        if (FAILED(existingHr) && ! IsDriveNotFound(existingHr))
        {
            return existingHr;
        }
        if (SUCCEEDED(existingHr) && (existing.isFolder || ! allowOverwrite))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        wil::unique_hfile tempFile;
        hr = Common::Files::CreateDeleteOnCloseTemporaryFile(kDriveTemporaryFileOptions, tempFile);
        if (FAILED(hr))
        {
            return hr;
        }
        auto* instance = new (std::nothrow) GoogleDriveFileWriter(this,
                                                                  std::move(connection),
                                                                  std::move(parent.id),
                                                                  std::move(leafName),
                                                                  allowOverwrite,
                                                                  std::move(tempFile),
                                                                  SUCCEEDED(existingHr) ? std::move(existing) : GoogleItem{});
        if (! instance)
        {
            return E_OUTOFMEMORY;
        }
        *writer = instance;
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive CreateFileWriter failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept
{
    if (! path || ! info)
    {
        return E_POINTER;
    }
    if (info->sizeBytes < sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }
    try
    {
        ResolvedConnection connection;
        HRESULT hr = ResolveConnection(path, true, connection);
        if (FAILED(hr))
        {
            return hr;
        }
        GoogleItem item;
        hr = ResolveItemByPath(connection, connection.canonicalPath, item);
        if (FAILED(hr))
        {
            return hr;
        }
        info->creationTime   = item.creationTime;
        info->lastAccessTime = item.lastAccessTime;
        info->lastWriteTime  = item.lastWriteTime;
        info->attributes     = item.attributes;
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive GetFileBasicInformation failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept
{
    if (! path || ! info)
    {
        return E_POINTER;
    }
    if (info->sizeBytes < sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }
    try
    {
        // Drive keeps one writable timestamp (modifiedTime); attributes have no counterpart.
        if (info->lastWriteTime == 0)
        {
            return S_OK;
        }
        const std::string modifiedTime = Rfc3339FromFileTime64(info->lastWriteTime);
        if (modifiedTime.empty())
        {
            return E_INVALIDARG;
        }
        ResolvedConnection connection;
        HRESULT hr = ResolveConnection(path, true, connection);
        if (FAILED(hr))
        {
            return hr;
        }
        if (connection.readOnly)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        GoogleItem item;
        hr = ResolveItemByPath(connection, connection.canonicalPath, item);
        if (FAILED(hr))
        {
            return hr;
        }
        std::string url = FilesEndpointUrl() + "/" + UrlEncodeUtf8(Utf8FromUtf16(item.id));
        AppendQueryParam(url, "supportsAllDrives", "true");
        AppendQueryParam(url, "fields", std::string(kDriveItemFields));
        const std::vector<std::string> headers{"Content-Type: application/json; charset=UTF-8"};
        long statusCode = 0;
        std::string responseBody;
        hr = PerformAuthorizedRequest(connection,
                                      "PATCH",
                                      url,
                                      headers,
                                      std::format(R"({{"modifiedTime":"{}"}})", modifiedTime),
                                      kMaxJsonResponseBytes,
                                      statusCode,
                                      responseBody,
                                      nullptr,
                                      nullptr,
                                      AuthorizedRequestRetry::Mutation);
        if (FAILED(hr))
        {
            return hr;
        }
        return MapHttpStatusToHresult(statusCode);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive SetFileBasicInformation failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept
{
    if (! jsonUtf8)
    {
        return E_POINTER;
    }
    *jsonUtf8 = nullptr;
    if (! path)
    {
        return E_POINTER;
    }
    try
    {
        ResolvedConnection connection;
        HRESULT hr = ResolveConnection(path, true, connection);
        if (FAILED(hr))
        {
            return hr;
        }
        GoogleItem item;
        hr = ResolveItemByPath(connection, connection.canonicalPath, item);
        if (FAILED(hr))
        {
            return hr;
        }
        std::string json = std::format(R"({{"provider":"google-drive","id":{},"name":{},"mimeType":{},"version":{},"size":{},"isFolder":{},"isShortcut":{}}})",
                                       JsonQuote(item.id),
                                       JsonQuote(item.name),
                                       JsonQuote(item.mimeType),
                                       JsonQuote(item.version),
                                       item.sizeBytes,
                                       item.isFolder ? "true" : "false",
                                       item.isShortcut ? "true" : "false");
        std::scoped_lock lock(_stateMutex);
        _itemPropertiesJson = std::move(json);
        *jsonUtf8           = _itemPropertiesJson.c_str();
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive GetItemProperties failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::CreateDirectory(const wchar_t* path) noexcept
{
    if (! path)
    {
        return E_POINTER;
    }
    try
    {
        ResolvedConnection connection;
        HRESULT hr = ResolveConnection(path, true, connection);
        if (FAILED(hr))
        {
            return hr;
        }
        if (connection.readOnly)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        if (connection.canonicalPath == L"/")
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        GoogleItem parent;
        std::wstring leafName;
        hr = ResolveParentAndLeaf(connection, connection.canonicalPath, parent, leafName);
        if (FAILED(hr))
        {
            return hr;
        }
        if (! IsValidDriveLeafName(leafName))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
        }
        GoogleItem existing;
        const HRESULT existingHr = ResolveChildByName(connection, parent.id, leafName, existing);
        if (SUCCEEDED(existingHr))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        if (! IsDriveNotFound(existingHr))
        {
            return existingHr;
        }
        GoogleItem created;
        return CreateFolderItem(connection, parent.id, leafName, created);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive CreateDirectory failed with std::exception.");
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetDirectorySize(
    const wchar_t* path, FileSystemFlags flags, IFileSystemDirectorySizeCallback* callback, void* cookie, FileSystemDirectorySizeResult* result) noexcept
{
    if (! path || ! result)
    {
        return E_POINTER;
    }
    if (result->sizeBytes < sizeof(FileSystemDirectorySizeResult))
    {
        return E_INVALIDARG;
    }
    result->totalBytes     = 0u;
    result->fileCount      = 0u;
    result->directoryCount = 0u;
    result->status         = S_OK;
    try
    {
        ResolvedConnection connection;
        HRESULT hr = ResolveConnection(path, true, connection);
        if (FAILED(hr))
        {
            result->status = hr;
            return hr;
        }
        GoogleItem item;
        hr = ResolveItemByPath(connection, connection.canonicalPath, item);
        if (FAILED(hr))
        {
            result->status = hr;
            return hr;
        }
        if (! item.isFolder)
        {
            result->totalBytes = item.sizeBytes;
            result->fileCount  = 1u;
            return S_OK;
        }

        struct PendingFolder final
        {
            std::wstring id;
            std::wstring path;
        };
        const bool recursive = (flags & FILESYSTEM_FLAG_RECURSIVE) != 0;
        std::vector<PendingFolder> pending;
        pending.push_back(PendingFolder{.id = item.id, .path = connection.canonicalPath});
        uint64_t scannedEntries = 0u;
        while (! pending.empty())
        {
            PendingFolder current = std::move(pending.back());
            pending.pop_back();
            if (callback)
            {
                BOOL cancel = FALSE;
                if (SUCCEEDED(callback->DirectorySizeShouldCancel(&cancel, cookie)) && cancel)
                {
                    result->status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    return result->status;
                }
            }
            std::vector<GoogleItem> children;
            hr = ListChildren(connection, current.id, children);
            if (FAILED(hr))
            {
                result->status = hr;
                return hr;
            }
            for (GoogleItem& child : children)
            {
                ++scannedEntries;
                if (child.isFolder)
                {
                    ++result->directoryCount;
                    if (recursive)
                    {
                        pending.push_back(PendingFolder{.id = std::move(child.id), .path = JoinDrivePath(current.path, child.name)});
                    }
                }
                else
                {
                    ++result->fileCount;
                    result->totalBytes += child.sizeBytes;
                }
            }
            if (callback)
            {
                static_cast<void>(callback->DirectorySizeProgress(
                    scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, current.path.c_str(), cookie));
            }
        }
        return S_OK;
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        Debug::Error(L"Google Drive GetDirectorySize failed with std::exception.");
        result->status = E_FAIL;
        return E_FAIL;
    }
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::SupportsAtomicWriterCommit(const wchar_t* path, FileSystemFlags /*flags*/, BOOL* supported) noexcept
{
    if (! supported)
    {
        return E_POINTER;
    }
    *supported = FALSE;
    if (! path)
    {
        return E_POINTER;
    }
    // A resumable upload publishes the destination name only when Drive acknowledges the final
    // chunk, and an overwrite replaces the existing object's content under its own id.
    *supported = TRUE;
    return S_OK;
}

unsigned long FileSystemGoogleDriveInternal::DriveProviderWatchdogTimeoutMs(uint32_t connectTimeoutMs, uint32_t requestTimeoutMs) noexcept
{
    // A silent server: one token request plus one API request, each bounded by the TCP connect
    // timeout and the hard request timeout (transport failures are not retried), plus 1 s of slack.
    const uint64_t perRequestMs = static_cast<uint64_t>(connectTimeoutMs) + static_cast<uint64_t>(requestTimeoutMs);
    const uint64_t totalMs      = perRequestMs * 2u + 1'000u;
    return static_cast<unsigned long>((std::min<uint64_t>)(totalMs, (std::numeric_limits<unsigned long>::max)()));
}

#if defined(ENABLE_TESTS)
void FileSystemGoogleDriveSelfTest::ConfigureFakeDrive(const wchar_t* origin, bool enable) noexcept
{
    {
        std::lock_guard lock(g_debugDriveOriginMutex);
        try
        {
            g_debugDriveOrigin = enable && origin != nullptr ? Utf8FromUtf16(origin) : std::string{};
        }
        catch (const std::bad_alloc&)
        {
            std::terminate();
        }
        catch (const std::exception&)
        {
            Debug::Error(L"Google Drive fake endpoint configuration failed with std::exception.");
            g_debugDriveOrigin.clear();
        }
    }
    g_debugDriveSyntheticConnection.store(enable, std::memory_order_release);
}
#endif
