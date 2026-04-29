#include "FileSystemGoogleDrive.h"

#include <algorithm>
#include <charconv>
#include <cstring>
#include <format>
#include <limits>
#include <map>
#include <memory>
#include <new>
#include <optional>
#include <utility>

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include <curl/curl.h>

#include "FileSystemGoogleDriveResources.h"
#include "Helpers.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

namespace
{
constexpr wchar_t kPluginId[]      = L"builtin/file-system-gdrive";
constexpr wchar_t kPluginShortId[] = L"gdrive";
constexpr wchar_t kPluginAuthor[]  = L"RedSalamander";
constexpr wchar_t kPluginVersion[] = VERSINFO_PLUGIN_VERSION;

[[nodiscard]] const wchar_t* LocalizedPluginName() noexcept
{
    static const std::wstring name = LoadStringResource(g_hInstance, IDS_FILESYSTEMGOOGLEDRIVE_NAME);
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

constexpr char kFolderMimeType[]   = "application/vnd.google-apps.folder";
constexpr char kShortcutMimeType[] = "application/vnd.google-apps.shortcut";

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
using unique_yyjson_doc = wil::unique_any<yyjson_doc*, decltype(&yyjson_doc_free), yyjson_doc_free>;

struct HttpResponse
{
    long statusCode = 0;
    std::string body;
};

struct LessNoCase
{
    bool operator()(const std::wstring& a, const std::wstring& b) const noexcept
    {
        return OrdinalString::LessNoCase(a, b);
    }
};

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

[[nodiscard]] std::optional<std::wstring> TryGetJsonString(yyjson_val* root, const char* key) noexcept
{
    if (! root || ! yyjson_is_obj(root) || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* value = yyjson_obj_get(root, key);
    if (! value || ! yyjson_is_str(value))
    {
        return std::nullopt;
    }

    const char* text    = yyjson_get_str(value);
    const size_t length = yyjson_get_len(value);
    if (! text)
    {
        return std::nullopt;
    }

    return Utf16FromUtf8(std::string_view(text, length));
}

[[nodiscard]] std::optional<std::string> TryGetJsonUtf8String(yyjson_val* root, const char* key) noexcept
{
    if (! root || ! yyjson_is_obj(root) || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* value = yyjson_obj_get(root, key);
    if (! value || ! yyjson_is_str(value))
    {
        return std::nullopt;
    }

    const char* text    = yyjson_get_str(value);
    const size_t length = yyjson_get_len(value);
    if (! text)
    {
        return std::nullopt;
    }

    return std::string(text, length);
}

[[nodiscard]] std::optional<bool> TryGetJsonBool(yyjson_val* root, const char* key) noexcept
{
    if (! root || ! yyjson_is_obj(root) || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* value = yyjson_obj_get(root, key);
    if (! value || ! yyjson_is_bool(value))
    {
        return std::nullopt;
    }

    return yyjson_get_bool(value) != 0;
}

[[nodiscard]] std::optional<uint64_t> TryGetJsonUInt(yyjson_val* root, const char* key) noexcept
{
    if (! root || ! yyjson_is_obj(root) || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* value = yyjson_obj_get(root, key);
    if (! value || ! yyjson_is_uint(value))
    {
        return std::nullopt;
    }

    return yyjson_get_uint(value);
}

[[nodiscard]] bool TryParseUInt64(std::string_view text, uint64_t& value) noexcept
{
    value = 0;
    if (text.empty())
    {
        return false;
    }

    const char* begin = text.data();
    const char* end   = begin + text.size();
    const auto result = std::from_chars(begin, end, value);
    return result.ec == std::errc{} && result.ptr == end;
}

[[nodiscard]] std::optional<uint64_t> TryGetJsonUInt64Flexible(yyjson_val* root, const char* key) noexcept
{
    if (! root || ! yyjson_is_obj(root) || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* value = yyjson_obj_get(root, key);
    if (! value)
    {
        return std::nullopt;
    }

    if (yyjson_is_uint(value))
    {
        return yyjson_get_uint(value);
    }

    if (yyjson_is_sint(value))
    {
        const int64_t signedValue = yyjson_get_sint(value);
        if (signedValue < 0)
        {
            return std::nullopt;
        }
        return static_cast<uint64_t>(signedValue);
    }

    if (yyjson_is_str(value))
    {
        const char* text    = yyjson_get_str(value);
        const size_t length = yyjson_get_len(value);
        uint64_t parsed     = 0;
        if (text && TryParseUInt64(std::string_view(text, length), parsed))
        {
            return parsed;
        }
    }

    return std::nullopt;
}

[[nodiscard]] bool IsUnreservedUrlByte(unsigned char ch) noexcept
{
    return (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9') || ch == '-' || ch == '.' || ch == '_' || ch == '~';
}

[[nodiscard]] std::string UrlEncodeUtf8(std::string_view text)
{
    std::string encoded;
    encoded.reserve(text.size() * 3u);

    constexpr char kHex[] = "0123456789ABCDEF";
    for (const char chValue : text)
    {
        const unsigned char ch = static_cast<unsigned char>(chValue);
        if (IsUnreservedUrlByte(ch))
        {
            encoded.push_back(static_cast<char>(ch));
            continue;
        }

        encoded.push_back('%');
        encoded.push_back(kHex[(ch >> 4u) & 0x0Fu]);
        encoded.push_back(kHex[ch & 0x0Fu]);
    }

    return encoded;
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
    static std::once_flag flag;
    static HRESULT initHr = E_FAIL;

    std::call_once(flag,
                   []() noexcept
    {
        const CURLcode code = curl_global_init(CURL_GLOBAL_DEFAULT);
        initHr              = (code == CURLE_OK) ? S_OK : E_FAIL;
    });

    return initHr;
}

size_t CurlWriteToString(char* data, size_t size, size_t nmemb, void* userData) noexcept
{
    const size_t bytes = size * nmemb;
    if (bytes == 0u || ! userData)
    {
        return bytes;
    }

    auto* out = static_cast<std::string*>(userData);
    // Mandatory C callback boundary: curl must not see C++ exceptions.
    try
    {
        out->append(data, bytes);
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
                           HttpResponse& response) noexcept
{
    response = {};

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
    curl_easy_setopt(curl.get(), CURLOPT_URL, urlCopy.c_str());
    curl_easy_setopt(curl.get(), CURLOPT_NOSIGNAL, 1L);
    curl_easy_setopt(curl.get(), CURLOPT_ACCEPT_ENCODING, "");
    curl_easy_setopt(curl.get(), CURLOPT_HTTP_VERSION, CURL_HTTP_VERSION_2TLS);
    curl_easy_setopt(curl.get(), CURLOPT_USERAGENT, "RedSalamander/GoogleDrive/0.1");
    curl_easy_setopt(curl.get(), CURLOPT_WRITEFUNCTION, &CurlWriteToString);
    curl_easy_setopt(curl.get(), CURLOPT_WRITEDATA, &response.body);

    if (connectTimeoutMs > 0)
    {
        curl_easy_setopt(curl.get(), CURLOPT_CONNECTTIMEOUT_MS, static_cast<long>(connectTimeoutMs));
    }

    if (requestTimeoutMs > 0)
    {
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
    if (name.empty())
    {
        return std::format(L"[id:{}]", id);
    }

    return std::format(L"{} [id:{}]", name, id);
}

[[nodiscard]] std::optional<std::pair<std::wstring, std::wstring>> TryParseSyntheticId(std::wstring_view displayName)
{
    constexpr std::wstring_view kMarker = L" [id:";
    if (! displayName.ends_with(L"]"))
    {
        return std::nullopt;
    }

    const size_t markerPos = displayName.rfind(kMarker);
    if (markerPos == std::wstring_view::npos)
    {
        return std::nullopt;
    }

    const size_t idStart = markerPos + kMarker.size();
    const size_t idEnd   = displayName.size() - 1u;
    if (idStart >= idEnd)
    {
        return std::nullopt;
    }

    return std::make_pair(std::wstring(displayName.substr(0, markerPos)), std::wstring(displayName.substr(idStart, idEnd - idStart)));
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

FileSystemGoogleDrive::~FileSystemGoogleDrive() = default;

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
    std::unique_lock lock(_stateMutex);
    ++_navigationMenuCallbackGeneration;
    _navigationMenuCallback       = callback;
    _navigationMenuCallbackCookie = callback ? cookie : nullptr;
    if (! callback)
    {
        _navigationMenuDrainCv.wait(lock, [this]() noexcept { return _navigationMenuCallbacksInFlight == 0; });
    }
    return S_OK;
}

bool FileSystemGoogleDrive::TryCaptureNavigationMenuCallback(NavigationMenuCallbackSnapshot& snapshot) noexcept
{
    std::scoped_lock lock(_stateMutex);
    if (! _navigationMenuCallback)
    {
        snapshot = {};
        return false;
    }

    snapshot.callback   = _navigationMenuCallback;
    snapshot.cookie     = _navigationMenuCallbackCookie;
    snapshot.generation = _navigationMenuCallbackGeneration;
    return true;
}

HRESULT FileSystemGoogleDrive::InvokeNavigationMenuCallback(const NavigationMenuCallbackSnapshot& snapshot, const wchar_t* path) noexcept
{
    INavigationMenuCallback* callback = nullptr;
    void* cookie                      = nullptr;
    {
        std::unique_lock lock(_stateMutex);
        if (_navigationMenuCallbackGeneration != snapshot.generation || _navigationMenuCallback != snapshot.callback)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        ++_navigationMenuCallbacksInFlight;
        callback = snapshot.callback;
        cookie   = snapshot.cookie;
    }

    const HRESULT callbackHr = callback->NavigationMenuRequestNavigate(path, cookie);

    {
        std::scoped_lock lock(_stateMutex);
        --_navigationMenuCallbacksInFlight;
    }
    _navigationMenuDrainCv.notify_all();
    return callbackHr;
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
        configuration = configurationJsonUtf8;
        yyjson_read_err err{};
        unique_yyjson_doc doc(yyjson_read_opts(configuration.data(), configuration.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err));
        if (doc)
        {
            yyjson_val* root = yyjson_doc_get_root(doc.get());
            if (root && yyjson_is_obj(root))
            {
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
    if (connection.clientId.empty() || connection.refreshToken.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_AUTHENTICATED);
    }

    const uint64_t now = GetTickCount64();
    {
        std::scoped_lock lock(_tokenMutex);
        const auto it = _accessTokensByConnectionKey.find(connection.connectionKey);
        if (it != _accessTokensByConnectionKey.end() && ! it->second.token.empty() && it->second.expiresAtTickMs > now + 30'000ull)
        {
            accessToken = it->second.token;
            return S_OK;
        }
    }

    const std::string clientIdUtf8     = Utf8FromUtf16(connection.clientId);
    const std::string refreshTokenUtf8 = Utf8FromUtf16(connection.refreshToken);
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

    const std::vector<std::string> headers = {
        "Content-Type: application/x-www-form-urlencoded",
        "Accept: application/json",
    };

    HttpResponse response{};
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

    uint64_t expiresIn = 3600;
    if (const auto value = TryGetJsonUInt64Flexible(root, "expires_in"); value.has_value() && value.value() > 0)
    {
        expiresIn = value.value();
    }

    accessToken = token.value();

    AccessTokenCacheEntry cacheEntry{};
    cacheEntry.token           = accessToken;
    cacheEntry.expiresAtTickMs = (expiresIn > 60u) ? (now + ((expiresIn - 60u) * 1000ull)) : (now + 30'000ull);

    {
        std::scoped_lock lock(_tokenMutex);
        _accessTokensByConnectionKey[connection.connectionKey] = std::move(cacheEntry);
    }

    return S_OK;
}

HRESULT FileSystemGoogleDrive::PerformAuthorizedJsonGet(const ResolvedConnection& connection, std::string_view url, std::string& body)
{
    for (int attempt = 0; attempt < 2; ++attempt)
    {
        std::wstring accessToken;
        HRESULT hr = GetAccessToken(connection, accessToken);
        if (FAILED(hr))
        {
            return hr;
        }

        const std::string accessTokenUtf8 = Utf8FromUtf16(accessToken);
        if (accessTokenUtf8.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        }

        const std::vector<std::string> headers = {
            std::format("Authorization: Bearer {}", accessTokenUtf8),
            "Accept: application/json",
        };

        HttpResponse response{};
        hr = PerformHttpRequest("GET", url, headers, {}, connection.connectTimeoutMs, connection.requestTimeoutMs, response);
        if (FAILED(hr))
        {
            return hr;
        }

        if (response.statusCode == 401 && attempt == 0)
        {
            std::scoped_lock lock(_tokenMutex);
            _accessTokensByConnectionKey.erase(connection.connectionKey);
            continue;
        }

        if (response.statusCode < 200 || response.statusCode >= 300)
        {
            return MapHttpStatusToHresult(response.statusCode);
        }

        body = std::move(response.body);
        return S_OK;
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_AUTHENTICATED);
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
    do
    {
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

        yyjson_val* files = yyjson_obj_get(root, "files");
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

        const auto explicitId = TryParseSyntheticId(segment);
        GoogleItem* match     = nullptr;
        bool duplicateName    = false;

        for (GoogleItem& child : children)
        {
            if (explicitId.has_value())
            {
                if (OrdinalString::EqualsNoCase(child.name, explicitId->first) && OrdinalString::EqualsNoCase(child.id, explicitId->second))
                {
                    match = &child;
                    break;
                }
                continue;
            }

            if (! OrdinalString::EqualsNoCase(child.name, segment))
            {
                continue;
            }

            if (! match)
            {
                match = &child;
            }
            else
            {
                duplicateName = true;
            }
        }

        if (! explicitId.has_value() && duplicateName)
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

    std::map<std::wstring, size_t, LessNoCase> counts;
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
        entry.name         = (countIt != counts.end() && countIt->second > 1u) ? MakeSyntheticDisplayName(child.name, child.id) : child.name;
        if (entry.name.empty())
        {
            entry.name = MakeSyntheticDisplayName(L"(unnamed)", child.id);
        }

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
