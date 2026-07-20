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

#include "CurlProcessRuntime.h"
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

constexpr char kTokenEndpoint[]  = "https://oauth2.googleapis.com/token";
constexpr char kFilesEndpoint[]  = "https://www.googleapis.com/drive/v3/files";
constexpr char kAboutEndpoint[]  = "https://www.googleapis.com/drive/v3/about";
constexpr char kDrivesEndpoint[] = "https://www.googleapis.com/drive/v3/drives/";

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
  "version": 1,
  "operations": {
    "copy": false,
    "move": false,
    "delete": false,
    "rename": false,
    "properties": false,
    "read": false,
    "write": false
  },
  "concurrency": {
    "copyMoveMax": 1,
    "deleteMax": 1,
    "deleteRecycleBinMax": 1
  },
  "crossFileSystem": {
    "export": { "copy": [], "move": [] },
    "import": { "copy": [], "move": [] }
  },
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": false,
    "componentComparison": "ordinalCaseSensitive",
    "normalization": "none",
    "preferredSeparator": "/",
    "acceptedSeparators": ["/"],
    "casePreserving": true,
    "caseOnlyRename": "notApplicable"
  }
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
};

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
    constexpr std::string_view kRetryAfter = "Retry-After:";
    if (line.size() < kRetryAfter.size() || ! std::equal(kRetryAfter.begin(), kRetryAfter.end(), line.begin(), [](char left, char right) noexcept {
        return static_cast<char>(std::tolower(static_cast<unsigned char>(left))) == static_cast<char>(std::tolower(static_cast<unsigned char>(right)));
    }))
    {
        return bytes;
    }

    line.remove_prefix(kRetryAfter.size());
    while (! line.empty() && (line.front() == ' ' || line.front() == '\t'))
    {
        line.remove_prefix(1u);
    }
    while (! line.empty() && (line.back() == '\r' || line.back() == '\n' || line.back() == ' ' || line.back() == '\t'))
    {
        line.remove_suffix(1u);
    }

    auto* response = static_cast<HttpResponse*>(userData);
    try
    {
        response->retryAfter.assign(line);
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
    uint64_t deadlineTickMs = 0u;
    HRESULT abortStatus     = S_OK;
};

int CurlCheckTransferDeadline(void* userData, curl_off_t, curl_off_t, curl_off_t, curl_off_t) noexcept
{
    auto* deadline = static_cast<CurlTransferDeadline*>(userData);
    if (deadline && deadline->deadlineTickMs != 0u && GetTickCount64() >= deadline->deadlineTickMs)
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

    const bool isPost = method == "POST";
    std::string bodyCopy;
    if (isPost)
    {
        bodyCopy.assign(body);
        curl_easy_setopt(curl.get(), CURLOPT_POST, 1L);
        curl_easy_setopt(curl.get(), CURLOPT_POSTFIELDS, bodyCopy.c_str());
        curl_easy_setopt(curl.get(), CURLOPT_POSTFIELDSIZE_LARGE, static_cast<curl_off_t>(bodyCopy.size()));
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
    unsigned long attributes = FILE_ATTRIBUTE_NORMAL;
    uint64_t sizeBytes       = 0;
    __int64 creationTime     = 0;
    __int64 lastAccessTime   = 0;
    __int64 lastWriteTime    = 0;
    __int64 changeTime       = 0;
    bool isFolder            = false;
    bool isShortcut          = false;
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
    request.version      = 1;
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
        request.version        = 1;
        request.sizeBytes      = sizeof(request);
        request.filterPluginId = kPluginId;
        request.ownerWindow    = nullptr;

        HostConnectionManagerResult result{};
        result.version   = 1;
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

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::CopyItem(const wchar_t* /*sourcePath*/,
                                                          const wchar_t* /*destinationPath*/,
                                                          FileSystemFlags /*flags*/,
                                                          const FileSystemOptions* /*options*/,
                                                          IFileSystemCallback* /*callback*/,
                                                          void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::MoveItem(const wchar_t* /*sourcePath*/,
                                                          const wchar_t* /*destinationPath*/,
                                                          FileSystemFlags /*flags*/,
                                                          const FileSystemOptions* /*options*/,
                                                          IFileSystemCallback* /*callback*/,
                                                          void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::DeleteItem(
    const wchar_t* /*path*/, FileSystemFlags /*flags*/, const FileSystemOptions* /*options*/, IFileSystemCallback* /*callback*/, void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::RenameItem(const wchar_t* /*sourcePath*/,
                                                            const wchar_t* /*destinationPath*/,
                                                            FileSystemFlags /*flags*/,
                                                            const FileSystemOptions* /*options*/,
                                                            IFileSystemCallback* /*callback*/,
                                                            void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::CopyItems(const wchar_t* const* /*sourcePaths*/,
                                                           unsigned long /*count*/,
                                                           const wchar_t* /*destinationFolder*/,
                                                           FileSystemFlags /*flags*/,
                                                           const FileSystemOptions* /*options*/,
                                                           IFileSystemCallback* /*callback*/,
                                                           void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::MoveItems(const wchar_t* const* /*sourcePaths*/,
                                                           unsigned long /*count*/,
                                                           const wchar_t* /*destinationFolder*/,
                                                           FileSystemFlags /*flags*/,
                                                           const FileSystemOptions* /*options*/,
                                                           IFileSystemCallback* /*callback*/,
                                                           void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::DeleteItems(const wchar_t* const* /*paths*/,
                                                             unsigned long /*count*/,
                                                             FileSystemFlags /*flags*/,
                                                             const FileSystemOptions* /*options*/,
                                                             IFileSystemCallback* /*callback*/,
                                                             void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::RenameItems(const FileSystemRenamePair* /*items*/,
                                                             unsigned long /*count*/,
                                                             FileSystemFlags /*flags*/,
                                                             const FileSystemOptions* /*options*/,
                                                             IFileSystemCallback* /*callback*/,
                                                             void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemGoogleDrive::GetCapabilities(const char** jsonUtf8) noexcept
{
    if (! jsonUtf8)
    {
        return E_POINTER;
    }

    *jsonUtf8 = kCapabilitiesJson;
    return S_OK;
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
        HRESULT hr = PerformHttpRequest("POST", kTokenEndpoint, headers, body, connection.connectTimeoutMs, connection.requestTimeoutMs, response);
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

HRESULT FileSystemGoogleDrive::PerformAuthorizedJsonGet(const ResolvedConnection& connection, std::string_view url, std::string& body)
{
    body.clear();
    const uint64_t operationDurationMs = std::clamp<uint64_t>(static_cast<uint64_t>(connection.requestTimeoutMs) * 2u, 30'000u, 600'000u);
    const uint64_t deadlineTickMs      = Common::Paging::DeadlineFromNow(GetTickCount64(), operationDurationMs);
    unsigned int retryCount            = 0u;
    bool refreshedAfterUnauthorized    = false;

    while (true)
    {
        const uint64_t nowTickMs = GetTickCount64();
        if (nowTickMs >= deadlineTickMs)
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

        std::vector<std::string> headers = {
            std::format("Authorization: Bearer {}", accessTokenUtf8),
            "Accept: application/json",
        };
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
        const uint64_t remainingMs      = deadlineTickMs - requestStartTickMs;
        const uint32_t requestTimeoutMs = static_cast<uint32_t>((std::min<uint64_t>)(remainingMs, (std::numeric_limits<uint32_t>::max)()));
        hr                              = PerformHttpRequest("GET", url, headers, {}, connection.connectTimeoutMs, requestTimeoutMs, response);
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

        if (IsRetryableAuthorizedStatus(response.statusCode) && retryCount < kMaxAuthorizedRetries)
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
                Sleep(static_cast<DWORD>(delayMs));
            }
            continue;
        }

        if (response.statusCode < 200 || response.statusCode >= 300)
        {
            return MapHttpStatusToHresult(response.statusCode);
        }

        body = std::move(response.body);
        return S_OK;
    }
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

        std::string url(kFilesEndpoint);
        AppendQueryParam(url, "fields", "nextPageToken,files(id,name,mimeType,modifiedTime,size,trashed)");
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

            if (const auto trashed = TryGetJsonBool(entry, "trashed"); trashed.value_or(false))
            {
                continue;
            }

            const auto idUtf8   = TryGetJsonUtf8String(entry, "id");
            const auto nameWide = TryGetJsonString(entry, "name");
            const auto mimeUtf8 = TryGetJsonUtf8String(entry, "mimeType");
            if (! idUtf8.has_value() || idUtf8->empty() || ! mimeUtf8.has_value())
            {
                continue;
            }

            GoogleItem item{};
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

        std::map<std::wstring, size_t, std::less<>> exactNameCounts;
        for (const GoogleItem& child : children)
        {
            ++exactNameCounts[child.name];
        }

        GoogleItem* match  = nullptr;
        bool duplicatePath = false;

        for (GoogleItem& child : children)
        {
            const auto count               = exactNameCounts.find(child.name);
            const std::wstring exposedName = MakeExposedItemName(child.name, child.id, count == exactNameCounts.end() ? 0u : count->second);
            if (exposedName != segment)
            {
                continue;
            }

            if (! match)
            {
                match = &child;
            }
            else
            {
                duplicatePath = true;
            }
        }

        if (duplicatePath)
        {
            return HRESULT_FROM_WIN32(ERROR_DUP_NAME);
        }
        if (! match)
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }

        item = *match;
    }

    return S_OK;
}

HRESULT FileSystemGoogleDrive::FetchDriveInfoPayload(const ResolvedConnection& connection, DriveInfoPayload& payload)
{
    payload = {};

    std::string aboutUrl(kAboutEndpoint);
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

    std::string driveUrl = std::string(kDrivesEndpoint) + UrlEncodeUtf8(Utf8FromUtf16(connection.sharedDriveId));
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
    return FileSystemGoogleDrive::RunDebugSelfTests(passed, failed);
}
#endif
