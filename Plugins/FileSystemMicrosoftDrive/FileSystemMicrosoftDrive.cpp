#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>

#include <bcrypt.h>

#include <shellapi.h>
#include <winhttp.h>

#include "FileSystemMicrosoftDrive.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cstring>
#include <format>
#include <limits>
#include <optional>
#include <span>
#include <string>
#include <string_view>
#include <thread>
#include <utility>

#pragma warning(push)
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include "Helpers.h"
#include "resource.h"
#include "FileSystemMicrosoftDriveResources.h"

extern HINSTANCE g_hInstance;

namespace FsMs
{
constexpr uint64_t kDefaultThrottleDelayMs    = 1'000ull;
constexpr uint64_t kAuthTimeoutMs             = 5ull * 60ull * 1'000ull;
constexpr uint64_t kGraphSimpleUploadMaxBytes = 250ull * 1024ull * 1024ull;
constexpr uint64_t kGraphChunkAlignmentBytes  = 320ull * 1024ull;

constexpr std::wstring_view kGraphBaseUrl         = L"https://graph.microsoft.com/v1.0";
constexpr std::wstring_view kAuthHost             = L"login.microsoftonline.com";
constexpr std::wstring_view kLoopbackPath         = L"/redsalamander/oauth2";
constexpr std::wstring_view kDefaultAuthUserAgent = L"RedSalamander Microsoft Drive/0.1";

constexpr std::wstring_view kScopeOneDrivePersonal = L"offline_access Files.ReadWrite User.Read openid profile";
constexpr std::wstring_view kScopeOneDriveBusiness = L"offline_access Files.ReadWrite User.Read openid profile";
constexpr std::wstring_view kScopeSharePoint       = L"offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read openid profile";

struct ConnectionProfileInfo
{
    std::wstring id;
    std::wstring name;
    std::wstring pluginId;
    std::wstring host;
    std::wstring initialPath = L"/";
    std::wstring userName;
    std::wstring authMode;
    std::wstring authorityOverride;
    std::wstring driveId;
    bool savePassword = false;
};

struct DriveContext
{
    ConnectionProfileInfo profile;
    std::wstring connectionName;
    std::wstring drivePath = L"/";
    std::wstring authority;
    std::wstring scopeText;
    std::wstring siteId;
    std::wstring driveId;
    std::wstring driveDisplayName;
    std::wstring driveVolumeLabel;
    std::wstring driveWebUrl;
    bool persistRefreshToken = false;
};

struct HttpHeader
{
    std::wstring_view name;
    std::wstring value;
};

struct HttpResponse
{
    DWORD statusCode = 0;
    std::string body;
    std::wstring retryAfter;
    std::wstring requestId;
    std::wstring location;
    std::wstring contentType;
};

struct ItemMetadata
{
    std::wstring id;
    std::wstring name;
    std::string rawJson;
    std::wstring downloadUrl;
    uint64_t sizeBytes       = 0;
    __int64 creationTime     = 0;
    __int64 lastAccessTime   = 0;
    __int64 lastWriteTime    = 0;
    __int64 changeTime       = 0;
    unsigned long attributes = 0;
    bool isFolder            = false;
};

struct AuthCodeResult
{
    std::wstring code;
    std::wstring state;
    std::wstring error;
    std::wstring errorDescription;
};

struct TokenResponse
{
    std::string accessToken;
    std::wstring refreshToken;
    uint64_t expiresInSeconds = 0;
    std::wstring error;
    std::wstring errorDescription;
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

[[nodiscard]] std::wstring LoadStringResourceOrFallback(unsigned int id, std::wstring_view fallback) noexcept
{
    std::wstring text = LoadStringResource(g_hInstance, id);
    if (! text.empty())
    {
        return text;
    }

    return std::wstring(fallback);
}

[[nodiscard]] const wchar_t* LocalizedPluginName(FileSystemMicrosoftDriveMode mode) noexcept
{
    switch (mode)
    {
        case FileSystemMicrosoftDriveMode::OneDrivePersonal:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_ONEDRIVE_PERSONAL_NAME);
            return text.c_str();
        }
        case FileSystemMicrosoftDriveMode::OneDriveBusiness:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_ONEDRIVE_BUSINESS_NAME);
            return text.c_str();
        }
        case FileSystemMicrosoftDriveMode::SharePoint:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_SHAREPOINT_NAME);
            return text.c_str();
        }
    }

    static const std::wstring fallback = LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_NAME);
    return fallback.c_str();
}

[[nodiscard]] const wchar_t* LocalizedPluginDescription(FileSystemMicrosoftDriveMode mode) noexcept
{
    switch (mode)
    {
        case FileSystemMicrosoftDriveMode::OneDrivePersonal:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_ONEDRIVE_PERSONAL_DESCRIPTION);
            return text.c_str();
        }
        case FileSystemMicrosoftDriveMode::OneDriveBusiness:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_ONEDRIVE_BUSINESS_DESCRIPTION);
            return text.c_str();
        }
        case FileSystemMicrosoftDriveMode::SharePoint:
        {
            static const std::wstring text = LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_SHAREPOINT_DESCRIPTION);
            return text.c_str();
        }
    }

    static const std::wstring fallback;
    return fallback.c_str();
}

[[nodiscard]] std::string HtmlEscapeUtf8(std::wstring_view text) noexcept
{
    const std::string utf8 = Utf8FromUtf16(text);
    if (utf8.empty() && ! text.empty())
    {
        return {};
    }

    std::string escaped;
    escaped.reserve(utf8.size());
    for (const char chValue : utf8)
    {
        const unsigned char ch = static_cast<unsigned char>(chValue);
        switch (ch)
        {
            case '&': escaped.append("&amp;"); break;
            case '<': escaped.append("&lt;"); break;
            case '>': escaped.append("&gt;"); break;
            case '"': escaped.append("&quot;"); break;
            case '\'': escaped.append("&#39;"); break;
            default: escaped.push_back(static_cast<char>(ch)); break;
        }
    }

    return escaped;
}

[[nodiscard]] std::string BuildAuthResultHttpResponse(bool success) noexcept
{
    const std::string appTitle = HtmlEscapeUtf8(LoadStringResourceOrFallback(IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_APP_TITLE, L"Red Salamander"));
    const std::string title    = HtmlEscapeUtf8(LoadStringResourceOrFallback(
        success ? IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_TITLE_SUCCESS : IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_TITLE_FAILURE, success ? L"You're connected" : L"We hit a snag"));
    const std::string body     = HtmlEscapeUtf8(
        LoadStringResourceOrFallback(success ? IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_SUCCESS : IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_FAILURE,
            success ? L"Microsoft sign-in is done. Head back to Red Salamander and keep exploring."
                    : L"Microsoft sign-in did not finish cleanly. Return to Red Salamander to review the problem and try again."));
    const std::string footer = HtmlEscapeUtf8(LoadStringResourceOrFallback(IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_FOOTER_HINT, L"This tab can be closed."));

    std::string html;
    html.reserve(2048);
    html.append("<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">");
    html.append("<title>");
    html.append(appTitle);
    html.append("</title><style>"
                "html,body{height:100%;margin:0;}"
                "body{display:grid;place-items:center;padding:24px;background:linear-gradient(180deg,#f7f1e6 0%,#efe5d6 100%);"
                "color:#241f1a;font-family:\"Segoe UI Variable Text\",\"Segoe UI\",Aptos,\"Noto Sans\",system-ui,sans-serif;}"
                ".card{width:min(100%,540px);background:#fffdf9;border:1px solid rgba(36,31,26,.10);border-radius:24px;"
                "box-shadow:0 22px 56px rgba(76,52,23,.12);padding:28px 28px 22px;}"
                ".icon{width:70px;height:70px;border-radius:20px;display:grid;place-items:center;font-size:34px;margin-bottom:18px;");
    html.append(success ? "background:linear-gradient(135deg,#d9f5e7,#fff7d9);" : "background:linear-gradient(135deg,#ffe0d8,#fff0d8);");
    html.append("}"
                ".title{margin:0 0 10px;font-size:28px;line-height:1.1;font-weight:700;letter-spacing:-.02em;}"
                ".body{margin:0;font-size:16px;line-height:1.55;color:#5c534b;}"
                ".footer{margin-top:18px;font-size:13px;color:#786f67;}"
                "</style></head><body><main class=\"card\"><div class=\"icon\" aria-hidden=\"true\">");
    html.append(success ? "&#x1F4C2;" : "&#x26A0;&#xFE0F;");
    html.append("</div><h1 class=\"title\">");
    html.append(title);
    html.append("</h1><p class=\"body\">");
    html.append(body);
    html.append("</p><p class=\"footer\">");
    html.append(footer);
    html.append("</p></main></body></html>");

    std::string response;
    response.reserve(html.size() + 256);
    response.append(success ? "HTTP/1.1 200 OK\r\n" : "HTTP/1.1 400 Bad Request\r\n");
    response.append("Content-Type: text/html; charset=utf-8\r\n");
    response.append("Cache-Control: no-store\r\n");
    response.append("Content-Security-Policy: default-src 'none'; style-src 'unsafe-inline'; img-src data:; base-uri 'none'; form-action 'none'\r\n");
    response.append("Connection: close\r\n\r\n");
    response.append(html);
    return response;
}

void SecureClear(std::string& text) noexcept
{
    if (! text.empty())
    {
        SecureZeroMemory(text.data(), text.size());
        text.clear();
    }
}

void SecureClear(std::wstring& text) noexcept
{
    if (! text.empty())
    {
        SecureZeroMemory(text.data(), text.size() * sizeof(wchar_t));
        text.clear();
    }
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

    if (! path.empty() && path.front() != L'/')
    {
        path.insert(path.begin(), L'/');
    }

    std::wstring collapsed;
    collapsed.reserve(path.size());

    bool prevSlash = false;
    for (const wchar_t ch : path)
    {
        const bool slash = ch == L'/';
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

[[nodiscard]] std::wstring CanonicalizeInputPath(std::wstring_view rawPath) noexcept
{
    std::wstring_view view(rawPath);
    const size_t colon = view.find(L':');
    if (colon != std::wstring_view::npos && colon > 0)
    {
        std::wstring_view suffix = view.substr(colon + 1u);
        if (! suffix.empty() && (suffix.front() == L'/' || suffix.front() == L'\\'))
        {
            return NormalizePluginPath(suffix);
        }
    }

    return NormalizePluginPath(view);
}

[[nodiscard]] std::wstring TrimTrailingSlashPreserveRoot(std::wstring path) noexcept
{
    while (path.size() > 1u && path.back() == L'/')
    {
        path.pop_back();
    }

    return path;
}

[[nodiscard]] std::vector<std::wstring_view> SplitPathSegments(std::wstring_view path) noexcept
{
    std::vector<std::wstring_view> segments;
    size_t index = 0;
    while (index < path.size())
    {
        while (index < path.size() && path[index] == L'/')
        {
            ++index;
        }

        if (index >= path.size())
        {
            break;
        }

        const size_t next = path.find(L'/', index);
        if (next == std::wstring_view::npos)
        {
            segments.push_back(path.substr(index));
            break;
        }

        segments.push_back(path.substr(index, next - index));
        index = next + 1u;
    }

    return segments;
}

[[nodiscard]] std::wstring JoinPath(std::wstring_view parent, std::wstring_view leaf) noexcept
{
    std::wstring result = NormalizePluginPath(parent);
    if (result.empty())
    {
        result = L"/";
    }

    if (result.size() > 1u && result.back() == L'/')
    {
        result.pop_back();
    }

    result.push_back(L'/');
    result.append(leaf);
    return NormalizePluginPath(result);
}

[[nodiscard]] HRESULT SplitParentAndLeaf(std::wstring_view path, std::wstring& parentOut, std::wstring& leafOut) noexcept
{
    std::wstring normalized = TrimTrailingSlashPreserveRoot(NormalizePluginPath(path));
    if (normalized == L"/" || normalized.empty())
    {
        return E_INVALIDARG;
    }

    const size_t slash = normalized.find_last_of(L'/');
    if (slash == std::wstring::npos)
    {
        return E_INVALIDARG;
    }

    parentOut = normalized.substr(0, slash);
    if (parentOut.empty())
    {
        parentOut = L"/";
    }

    leafOut = normalized.substr(slash + 1u);
    return leafOut.empty() ? E_INVALIDARG : S_OK;
}

[[nodiscard]] bool ParseFixedDigits(std::wstring_view text, size_t offset, size_t count, int& valueOut) noexcept
{
    if (offset + count > text.size())
    {
        return false;
    }

    int value = 0;
    for (size_t i = 0; i < count; ++i)
    {
        const wchar_t ch = text[offset + i];
        if (ch < L'0' || ch > L'9')
        {
            return false;
        }

        value = (value * 10) + static_cast<int>(ch - L'0');
    }

    valueOut = value;
    return true;
}

[[nodiscard]] __int64 ParseIso8601FileTime(std::wstring_view text) noexcept
{
    if (text.size() < 20u)
    {
        return 0;
    }

    int year   = 0;
    int month  = 0;
    int day    = 0;
    int hour   = 0;
    int minute = 0;
    int second = 0;
    if (! ParseFixedDigits(text, 0, 4, year) || text[4] != L'-' || ! ParseFixedDigits(text, 5, 2, month) || text[7] != L'-' ||
        ! ParseFixedDigits(text, 8, 2, day) || (text[10] != L'T' && text[10] != L't') || ! ParseFixedDigits(text, 11, 2, hour) || text[13] != L':' ||
        ! ParseFixedDigits(text, 14, 2, minute) || text[16] != L':' || ! ParseFixedDigits(text, 17, 2, second))
    {
        return 0;
    }

    SYSTEMTIME st{};
    st.wYear   = static_cast<WORD>(year);
    st.wMonth  = static_cast<WORD>(month);
    st.wDay    = static_cast<WORD>(day);
    st.wHour   = static_cast<WORD>(hour);
    st.wMinute = static_cast<WORD>(minute);
    st.wSecond = static_cast<WORD>(second);

    size_t index = 19u;
    if (index < text.size() && text[index] == L'.')
    {
        ++index;
        int milliseconds = 0;
        int factor       = 100;
        while (index < text.size())
        {
            const wchar_t ch = text[index];
            if (ch < L'0' || ch > L'9')
            {
                break;
            }

            if (factor > 0)
            {
                milliseconds += static_cast<int>(ch - L'0') * factor;
                factor /= 10;
            }

            ++index;
        }
        st.wMilliseconds = static_cast<WORD>(milliseconds);
    }

    if (index >= text.size() || (text[index] != L'Z' && text[index] != L'z'))
    {
        return 0;
    }

    FILETIME ft{};
    if (SystemTimeToFileTime(&st, &ft) == 0)
    {
        return 0;
    }

    ULARGE_INTEGER ull{};
    ull.LowPart  = ft.dwLowDateTime;
    ull.HighPart = ft.dwHighDateTime;
    return static_cast<__int64>(ull.QuadPart);
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

    const char* utf8 = yyjson_get_str(value);
    const size_t len = yyjson_get_len(value);
    if (! utf8)
    {
        return std::nullopt;
    }

    return Utf16FromUtf8(std::string_view(utf8, len));
}

[[nodiscard]] std::optional<uint64_t> TryGetJsonUInt(yyjson_val* root, const char* key) noexcept
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

    if (yyjson_is_int(value))
    {
        const int64_t signedValue = yyjson_get_sint(value);
        if (signedValue >= 0)
        {
            return static_cast<uint64_t>(signedValue);
        }
    }

    if (yyjson_is_str(value))
    {
        const char* text = yyjson_get_str(value);
        if (! text)
        {
            return std::nullopt;
        }

        uint64_t parsed = 0;
        for (const char* p = text; *p; ++p)
        {
            if (*p < '0' || *p > '9')
            {
                return std::nullopt;
            }

            if (parsed > ((std::numeric_limits<uint64_t>::max)() / 10ull))
            {
                return std::nullopt;
            }

            parsed = (parsed * 10ull) + static_cast<uint64_t>(*p - '0');
        }

        return parsed;
    }

    return std::nullopt;
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

[[nodiscard]] HRESULT GetTemporaryDeleteOnCloseFile(wil::unique_hfile& fileOut) noexcept
{
    fileOut.reset();

    wchar_t directory[MAX_PATH + 1] = {};
    const DWORD directoryLen        = GetTempPathW(static_cast<DWORD>(std::size(directory)), directory);
    if (directoryLen == 0 || directoryLen >= std::size(directory))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    wchar_t fileName[MAX_PATH + 1] = {};
    if (GetTempFileNameW(directory, L"rsm", 0, fileName) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    wil::unique_hfile temp(CreateFileW(
        fileName, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE, nullptr));
    if (! temp)
    {
        const DWORD lastError = GetLastError();
        DeleteFileW(fileName);
        return HRESULT_FROM_WIN32(lastError);
    }

    fileOut = std::move(temp);
    return S_OK;
}

[[nodiscard]] HRESULT ResetFilePointerToStart(HANDLE file) noexcept
{
    if (file == nullptr || file == INVALID_HANDLE_VALUE)
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

[[nodiscard]] HRESULT GetFileSizeBytes(HANDLE file, uint64_t& sizeBytesOut) noexcept
{
    sizeBytesOut = 0;

    if (file == nullptr || file == INVALID_HANDLE_VALUE)
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

    sizeBytesOut = static_cast<uint64_t>(size.QuadPart);
    return S_OK;
}

[[nodiscard]] HRESULT WriteBufferToFile(HANDLE file, const void* buffer, size_t bytesToWrite) noexcept
{
    if (file == nullptr || file == INVALID_HANDLE_VALUE)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    if (bytesToWrite == 0)
    {
        return S_OK;
    }

    if (! buffer)
    {
        return E_POINTER;
    }

    size_t offset = 0;
    while (offset < bytesToWrite)
    {
        const DWORD chunk = static_cast<DWORD>(std::min<size_t>(bytesToWrite - offset, static_cast<size_t>((std::numeric_limits<DWORD>::max)())));
        DWORD written     = 0;
        if (WriteFile(file, static_cast<const std::byte*>(buffer) + offset, chunk, &written, nullptr) == 0)
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

[[nodiscard]] std::wstring PercentEncodeUtf8(std::wstring_view text) noexcept
{
    const std::string utf8 = Utf8FromUtf16(text);
    if (utf8.empty() && ! text.empty())
    {
        return {};
    }

    std::wstring encoded;
    encoded.reserve(utf8.size());

    constexpr char kHex[] = "0123456789ABCDEF";
    for (const char chValue : utf8)
    {
        const unsigned char ch = static_cast<unsigned char>(chValue);
        const bool unreserved =
            (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9') || ch == '-' || ch == '.' || ch == '_' || ch == '~';
        if (unreserved)
        {
            encoded.push_back(static_cast<wchar_t>(ch));
            continue;
        }

        encoded.push_back(L'%');
        encoded.push_back(static_cast<wchar_t>(kHex[(ch >> 4) & 0x0F]));
        encoded.push_back(static_cast<wchar_t>(kHex[ch & 0x0F]));
    }

    return encoded;
}

[[nodiscard]] std::wstring PercentEncodeGraphPath(std::wstring_view path) noexcept
{
    const auto segments = SplitPathSegments(path);
    if (segments.empty())
    {
        return {};
    }

    std::wstring encoded;
    bool first = true;
    for (const std::wstring_view segment : segments)
    {
        if (! first)
        {
            encoded.push_back(L'/');
        }
        first = false;

        const std::wstring encodedSegment = PercentEncodeUtf8(segment);
        if (encodedSegment.empty() && ! segment.empty())
        {
            return {};
        }
        encoded.append(encodedSegment);
    }

    return encoded;
}

[[nodiscard]] bool TryAppendByteFromHex(wchar_t high, wchar_t low, std::string& out) noexcept
{
    auto decode = [](wchar_t ch) noexcept -> int
    {
        if (ch >= L'0' && ch <= L'9')
        {
            return static_cast<int>(ch - L'0');
        }
        if (ch >= L'A' && ch <= L'F')
        {
            return static_cast<int>(10 + (ch - L'A'));
        }
        if (ch >= L'a' && ch <= L'f')
        {
            return static_cast<int>(10 + (ch - L'a'));
        }
        return -1;
    };

    const int hi = decode(high);
    const int lo = decode(low);
    if (hi < 0 || lo < 0)
    {
        return false;
    }

    out.push_back(static_cast<char>((hi << 4) | lo));
    return true;
}

[[nodiscard]] std::wstring UrlDecodeToUtf16(std::wstring_view text) noexcept
{
    std::string utf8;
    utf8.reserve(text.size());

    for (size_t index = 0; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch == L'+')
        {
            utf8.push_back(' ');
            continue;
        }

        if (ch == L'%' && (index + 2u) < text.size())
        {
            if (TryAppendByteFromHex(text[index + 1u], text[index + 2u], utf8))
            {
                index += 2u;
                continue;
            }
        }

        if (ch > 0x7F)
        {
            return {};
        }

        utf8.push_back(static_cast<char>(ch));
    }

    return Utf16FromUtf8(utf8);
}

[[nodiscard]] std::wstring GetQueryValue(std::wstring_view query, std::wstring_view key) noexcept
{
    size_t index = 0;
    while (index < query.size())
    {
        const size_t amp                   = query.find(L'&', index);
        const std::wstring_view pair       = query.substr(index, amp == std::wstring_view::npos ? std::wstring_view::npos : (amp - index));
        const size_t equal                 = pair.find(L'=');
        const std::wstring_view currentKey = (equal == std::wstring_view::npos) ? pair : pair.substr(0, equal);
        if (OrdinalString::EqualsNoCase(currentKey, key))
        {
            const std::wstring_view value = (equal == std::wstring_view::npos) ? std::wstring_view{} : pair.substr(equal + 1u);
            return UrlDecodeToUtf16(value);
        }

        if (amp == std::wstring_view::npos)
        {
            break;
        }
        index = amp + 1u;
    }

    return {};
}

[[nodiscard]] std::wstring Base64UrlEncode(std::span<const std::byte> bytes) noexcept
{
    static constexpr char kAlphabet[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_";

    std::wstring out;
    out.reserve(((bytes.size() + 2u) / 3u) * 4u);

    size_t index = 0;
    while (index + 3u <= bytes.size())
    {
        const uint32_t block =
            (static_cast<uint32_t>(bytes[index]) << 16u) | (static_cast<uint32_t>(bytes[index + 1u]) << 8u) | static_cast<uint32_t>(bytes[index + 2u]);
        out.push_back(static_cast<wchar_t>(kAlphabet[(block >> 18u) & 0x3Fu]));
        out.push_back(static_cast<wchar_t>(kAlphabet[(block >> 12u) & 0x3Fu]));
        out.push_back(static_cast<wchar_t>(kAlphabet[(block >> 6u) & 0x3Fu]));
        out.push_back(static_cast<wchar_t>(kAlphabet[block & 0x3Fu]));
        index += 3u;
    }

    const size_t remaining = bytes.size() - index;
    if (remaining == 1u)
    {
        const uint32_t block = static_cast<uint32_t>(bytes[index]) << 16u;
        out.push_back(static_cast<wchar_t>(kAlphabet[(block >> 18u) & 0x3Fu]));
        out.push_back(static_cast<wchar_t>(kAlphabet[(block >> 12u) & 0x3Fu]));
    }
    else if (remaining == 2u)
    {
        const uint32_t block = (static_cast<uint32_t>(bytes[index]) << 16u) | (static_cast<uint32_t>(bytes[index + 1u]) << 8u);
        out.push_back(static_cast<wchar_t>(kAlphabet[(block >> 18u) & 0x3Fu]));
        out.push_back(static_cast<wchar_t>(kAlphabet[(block >> 12u) & 0x3Fu]));
        out.push_back(static_cast<wchar_t>(kAlphabet[(block >> 6u) & 0x3Fu]));
    }

    return out;
}

[[nodiscard]] HRESULT GenerateRandomBytes(std::span<std::byte> bytes) noexcept
{
    if (bytes.empty())
    {
        return S_OK;
    }

    const NTSTATUS status = BCryptGenRandom(nullptr, reinterpret_cast<PUCHAR>(bytes.data()), static_cast<ULONG>(bytes.size()), BCRYPT_USE_SYSTEM_PREFERRED_RNG);
    return BCRYPT_SUCCESS(status) ? S_OK : HRESULT_FROM_NT(status);
}

[[nodiscard]] HRESULT ComputeSha256(std::wstring_view text, std::array<std::byte, 32>& digestOut) noexcept
{
    const std::string utf8 = Utf8FromUtf16(text);
    if (utf8.empty() && ! text.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    wil::unique_bcrypt_algorithm algorithm;
    NTSTATUS status = BCryptOpenAlgorithmProvider(algorithm.put(), BCRYPT_SHA256_ALGORITHM, nullptr, 0);
    if (! BCRYPT_SUCCESS(status))
    {
        return HRESULT_FROM_NT(status);
    }

    DWORD hashObjectBytes = 0;
    DWORD resultBytes     = 0;
    status = BCryptGetProperty(algorithm.get(), BCRYPT_OBJECT_LENGTH, reinterpret_cast<PUCHAR>(&hashObjectBytes), sizeof(hashObjectBytes), &resultBytes, 0);
    if (! BCRYPT_SUCCESS(status) || hashObjectBytes == 0)
    {
        return HRESULT_FROM_NT(status);
    }

    std::vector<std::byte> hashObject(hashObjectBytes);
    wil::unique_bcrypt_hash hash;
    status = BCryptCreateHash(algorithm.get(), hash.put(), reinterpret_cast<PUCHAR>(hashObject.data()), static_cast<ULONG>(hashObject.size()), nullptr, 0, 0);
    if (! BCRYPT_SUCCESS(status))
    {
        return HRESULT_FROM_NT(status);
    }

    status = BCryptHashData(hash.get(), reinterpret_cast<PUCHAR>(const_cast<char*>(utf8.data())), static_cast<ULONG>(utf8.size()), 0);
    if (! BCRYPT_SUCCESS(status))
    {
        return HRESULT_FROM_NT(status);
    }

    status = BCryptFinishHash(hash.get(), reinterpret_cast<PUCHAR>(digestOut.data()), static_cast<ULONG>(digestOut.size()), 0);
    if (! BCRYPT_SUCCESS(status))
    {
        return HRESULT_FROM_NT(status);
    }

    return S_OK;
}

[[nodiscard]] HRESULT BuildPkceValues(std::wstring& verifierOut, std::wstring& challengeOut, std::wstring& stateOut) noexcept
{
    verifierOut.clear();
    challengeOut.clear();
    stateOut.clear();

    std::array<std::byte, 32> verifierBytes{};
    HRESULT hr = GenerateRandomBytes(verifierBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    std::array<std::byte, 16> stateBytes{};
    hr = GenerateRandomBytes(stateBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    verifierOut = Base64UrlEncode(verifierBytes);
    stateOut    = Base64UrlEncode(stateBytes);
    if (verifierOut.empty() || stateOut.empty())
    {
        return E_FAIL;
    }

    std::array<std::byte, 32> digest{};
    hr = ComputeSha256(verifierOut, digest);
    if (FAILED(hr))
    {
        SecureClear(verifierOut);
        SecureClear(stateOut);
        return hr;
    }

    challengeOut = Base64UrlEncode(digest);
    if (challengeOut.empty())
    {
        SecureClear(verifierOut);
        SecureClear(stateOut);
        return E_FAIL;
    }

    return S_OK;
}

[[nodiscard]] bool ParseRetryAfterMs(std::wstring_view retryAfter, uint64_t& delayMsOut) noexcept
{
    delayMsOut = 0;
    if (retryAfter.empty())
    {
        return false;
    }

    uint64_t seconds = 0;
    for (const wchar_t ch : retryAfter)
    {
        if (ch < L'0' || ch > L'9')
        {
            return false;
        }

        seconds = (seconds * 10ull) + static_cast<uint64_t>(ch - L'0');
    }

    delayMsOut = seconds * 1'000ull;
    return true;
}

[[nodiscard]] HRESULT EnsureWinsockStarted(WSADATA& dataOut) noexcept
{
    const int rc = WSAStartup(MAKEWORD(2, 2), &dataOut);
    return rc == 0 ? S_OK : HRESULT_FROM_WIN32(static_cast<unsigned long>(rc));
}

[[nodiscard]] HRESULT LastSocketErrorHresult() noexcept
{
    return HRESULT_FROM_WIN32(static_cast<unsigned long>(WSAGetLastError()));
}

[[nodiscard]] HRESULT SetExclusiveAddressUse(SOCKET socketHandle) noexcept
{
    BOOL exclusive = TRUE;
    if (setsockopt(socketHandle, SOL_SOCKET, SO_EXCLUSIVEADDRUSE, reinterpret_cast<const char*>(&exclusive), sizeof(exclusive)) != 0)
    {
        return LastSocketErrorHresult();
    }
    return S_OK;
}

[[nodiscard]] HRESULT CreateIpv4LoopbackListener(unsigned short requestedPort, wil::unique_socket& listenerOut, unsigned short& actualPortOut) noexcept
{
    listenerOut.reset();
    actualPortOut = 0;

    wil::unique_socket listener(socket(AF_INET, SOCK_STREAM, IPPROTO_TCP));
    if (! listener)
    {
        return LastSocketErrorHresult();
    }

    HRESULT hr = SetExclusiveAddressUse(listener.get());
    if (FAILED(hr))
    {
        return hr;
    }

    sockaddr_in address{};
    address.sin_family      = AF_INET;
    address.sin_addr.s_addr = htonl(INADDR_LOOPBACK);
    address.sin_port        = htons(requestedPort);
    if (bind(listener.get(), reinterpret_cast<const sockaddr*>(&address), sizeof(address)) != 0)
    {
        return LastSocketErrorHresult();
    }

    if (listen(listener.get(), 1) != 0)
    {
        return LastSocketErrorHresult();
    }

    sockaddr_in boundAddress{};
    int boundLength = sizeof(boundAddress);
    if (getsockname(listener.get(), reinterpret_cast<sockaddr*>(&boundAddress), &boundLength) != 0)
    {
        return LastSocketErrorHresult();
    }

    actualPortOut = ntohs(boundAddress.sin_port);
    listenerOut   = std::move(listener);
    return S_OK;
}

[[nodiscard]] HRESULT CreateIpv6LoopbackListener(unsigned short requestedPort, wil::unique_socket& listenerOut) noexcept
{
    listenerOut.reset();

    wil::unique_socket listener(socket(AF_INET6, SOCK_STREAM, IPPROTO_TCP));
    if (! listener)
    {
        return LastSocketErrorHresult();
    }

    HRESULT hr = SetExclusiveAddressUse(listener.get());
    if (FAILED(hr))
    {
        return hr;
    }

    sockaddr_in6 address{};
    address.sin6_family = AF_INET6;
    address.sin6_addr   = in6addr_loopback;
    address.sin6_port   = htons(requestedPort);
    if (bind(listener.get(), reinterpret_cast<const sockaddr*>(&address), sizeof(address)) != 0)
    {
        return LastSocketErrorHresult();
    }

    if (listen(listener.get(), 1) != 0)
    {
        return LastSocketErrorHresult();
    }

    listenerOut = std::move(listener);
    return S_OK;
}

[[nodiscard]] HRESULT WaitForAuthRedirect(std::wstring_view expectedPath,
                                          std::wstring_view expectedState,
                                          std::wstring& redirectUriOut,
                                          AuthCodeResult& resultOut) noexcept
{
    redirectUriOut.clear();
    resultOut = {};

    WSADATA wsa{};
    HRESULT hr = EnsureWinsockStarted(wsa);
    if (FAILED(hr))
    {
        return hr;
    }
    auto cleanupWsa = wil::scope_exit([&] { WSACleanup(); });

    wil::unique_socket listenerV4;
    unsigned short port = 0;
    hr                  = CreateIpv4LoopbackListener(0, listenerV4, port);
    if (FAILED(hr))
    {
        return hr;
    }

    wil::unique_socket listenerV6;
    const HRESULT ipv6Hr = CreateIpv6LoopbackListener(port, listenerV6);
    if (FAILED(ipv6Hr))
    {
        Debug::Warning(L"Microsoft Drive: IPv6 localhost auth listener unavailable (hr=0x{:08X}); falling back to IPv4 loopback only.",
                       static_cast<unsigned long>(ipv6Hr));
    }

    // Desktop app registrations use the localhost exception and vary only the loopback port at runtime.
    redirectUriOut = std::format(L"http://localhost:{}{}", port, expectedPath);

    fd_set readSet{};
    FD_ZERO(&readSet);
    FD_SET(listenerV4.get(), &readSet);
    if (listenerV6)
    {
        FD_SET(listenerV6.get(), &readSet);
    }

    timeval timeout{};
    timeout.tv_sec         = static_cast<long>(kAuthTimeoutMs / 1000ull);
    timeout.tv_usec        = static_cast<long>((kAuthTimeoutMs % 1000ull) * 1000ull);
    const int selectResult = select(0, &readSet, nullptr, nullptr, &timeout);
    if (selectResult == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
    }
    if (selectResult == SOCKET_ERROR)
    {
        return LastSocketErrorHresult();
    }

    wil::unique_socket client;
    if (listenerV6 && FD_ISSET(listenerV6.get(), &readSet))
    {
        client.reset(accept(listenerV6.get(), nullptr, nullptr));
    }
    else if (FD_ISSET(listenerV4.get(), &readSet))
    {
        client.reset(accept(listenerV4.get(), nullptr, nullptr));
    }

    if (! client)
    {
        return LastSocketErrorHresult();
    }

    std::array<char, 8 * 1024> requestBytes{};
    const int received = recv(client.get(), requestBytes.data(), static_cast<int>(requestBytes.size()) - 1, 0);
    if (received <= 0)
    {
        const unsigned long errorCode = received == 0 ? ERROR_CONNECTION_ABORTED : static_cast<unsigned long>(WSAGetLastError());
        return HRESULT_FROM_WIN32(errorCode);
    }

    requestBytes[static_cast<size_t>(received)] = '\0';
    const std::string_view requestView(requestBytes.data(), static_cast<size_t>(received));
    const size_t lineEnd = requestView.find("\r\n");
    if (lineEnd == std::string_view::npos)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const std::string_view requestLine = requestView.substr(0, lineEnd);
    if (! requestLine.starts_with("GET "))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const size_t pathStart = 4u;
    const size_t pathEnd   = requestLine.find(' ', pathStart);
    if (pathEnd == std::string_view::npos || pathEnd <= pathStart)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const std::wstring requestTarget = Utf16FromUtf8(requestLine.substr(pathStart, pathEnd - pathStart));
    if (requestTarget.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const size_t queryPos = requestTarget.find(L'?');
    const std::wstring_view requestPath =
        queryPos == std::wstring::npos ? std::wstring_view(requestTarget) : std::wstring_view(requestTarget).substr(0, queryPos);
    const std::wstring_view query = queryPos == std::wstring::npos ? std::wstring_view{} : std::wstring_view(requestTarget).substr(queryPos + 1u);
    const bool stateMatches       = OrdinalString::EqualsNoCase(requestPath, expectedPath);
    resultOut.code                = GetQueryValue(query, L"code");
    resultOut.state               = GetQueryValue(query, L"state");
    resultOut.error               = GetQueryValue(query, L"error");
    resultOut.errorDescription    = GetQueryValue(query, L"error_description");

    const bool valid = stateMatches && OrdinalString::EqualsNoCase(resultOut.state, expectedState) &&
                       ((! resultOut.code.empty() && resultOut.error.empty()) || (! resultOut.error.empty() && resultOut.code.empty()));

    const std::string response = BuildAuthResultHttpResponse(valid);
    static_cast<void>(send(client.get(), response.data(), static_cast<int>(response.size()), 0));
    shutdown(client.get(), SD_BOTH);

    return valid ? S_OK : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

[[nodiscard]] HINTERNET GetSharedWinHttpSession() noexcept
{
    struct SessionState
    {
        std::once_flag initOnce;
        wil::unique_winhttp_hinternet session;

        SessionState() = default;
        SessionState(const SessionState&) = delete;
        SessionState& operator=(const SessionState&) = delete;
        SessionState(SessionState&&) = delete;
        SessionState& operator=(SessionState&&) = delete;
    };

    static SessionState state;
    std::call_once(state.initOnce,
                   [&]() noexcept
    {
        state.session.reset(WinHttpOpen(kDefaultAuthUserAgent.data(), WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0));
        if (! state.session)
        {
            Debug::ErrorWithLastError(L"MicrosoftDrive: WinHttpOpen failed");
            return;
        }

        DWORD decompression = WINHTTP_DECOMPRESSION_FLAG_GZIP | WINHTTP_DECOMPRESSION_FLAG_DEFLATE;
        static_cast<void>(WinHttpSetOption(state.session.get(), WINHTTP_OPTION_DECOMPRESSION, &decompression, sizeof(decompression)));

#ifdef WINHTTP_OPTION_ENABLE_HTTP_PROTOCOL
        DWORD protocol = WINHTTP_PROTOCOL_FLAG_HTTP2;
        static_cast<void>(WinHttpSetOption(state.session.get(), WINHTTP_OPTION_ENABLE_HTTP_PROTOCOL, &protocol, sizeof(protocol)));
#endif
    });

    return state.session.get();
}

[[nodiscard]] std::wstring QueryStringHeader(HINTERNET request, DWORD query) noexcept
{
    DWORD required = 0;
    if (WinHttpQueryHeaders(request, query, WINHTTP_HEADER_NAME_BY_INDEX, nullptr, &required, WINHTTP_NO_HEADER_INDEX) == 0)
    {
        if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || required < sizeof(wchar_t))
        {
            return {};
        }
    }

    std::wstring value(required / sizeof(wchar_t), L'\0');
    if (WinHttpQueryHeaders(request, query, WINHTTP_HEADER_NAME_BY_INDEX, value.data(), &required, WINHTTP_NO_HEADER_INDEX) == 0)
    {
        return {};
    }

    if (! value.empty() && value.back() == L'\0')
    {
        value.pop_back();
    }
    return value;
}

[[nodiscard]] std::wstring QueryStringHeader(HINTERNET request, DWORD query, const wchar_t* headerName) noexcept
{
    DWORD required = 0;
    if (WinHttpQueryHeaders(request, query, headerName, nullptr, &required, WINHTTP_NO_HEADER_INDEX) == 0)
    {
        if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || required < sizeof(wchar_t))
        {
            return {};
        }
    }

    std::wstring value(required / sizeof(wchar_t), L'\0');
    if (WinHttpQueryHeaders(request, query, headerName, value.data(), &required, WINHTTP_NO_HEADER_INDEX) == 0)
    {
        return {};
    }

    if (! value.empty() && value.back() == L'\0')
    {
        value.pop_back();
    }
    return value;
}

[[nodiscard]] HRESULT CrackUrl(std::wstring_view url, URL_COMPONENTS& components, std::wstring& hostOut, std::wstring& pathAndQueryOut) noexcept
{
    hostOut.clear();
    pathAndQueryOut.clear();

    std::wstring urlCopy(url);
    components                   = {};
    components.dwStructSize      = sizeof(components);
    components.dwSchemeLength    = static_cast<DWORD>(-1);
    components.dwHostNameLength  = static_cast<DWORD>(-1);
    components.dwUrlPathLength   = static_cast<DWORD>(-1);
    components.dwExtraInfoLength = static_cast<DWORD>(-1);

    // We want pointers back into urlCopy; ICU_ESCAPE is only valid when WinHTTP copies into caller buffers.
    if (WinHttpCrackUrl(urlCopy.data(), static_cast<DWORD>(urlCopy.size()), 0, &components) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    hostOut.assign(components.lpszHostName, components.dwHostNameLength);
    pathAndQueryOut.assign(components.lpszUrlPath, components.dwUrlPathLength);
    if (components.dwExtraInfoLength > 0)
    {
        pathAndQueryOut.append(components.lpszExtraInfo, components.dwExtraInfoLength);
    }

    if (pathAndQueryOut.empty())
    {
        pathAndQueryOut = L"/";
    }

    return S_OK;
}

[[nodiscard]] HRESULT SendHttpRequest(const FileSystemMicrosoftDrive::Settings& settings,
                                      std::wstring_view method,
                                      std::wstring_view url,
                                      const char* bearerToken,
                                      std::span<const HttpHeader> headers,
                                      const std::byte* bodyBytes,
                                      size_t bodySizeBytes,
                                      HANDLE bodyFile,
                                      bool allowRetry,
                                      HttpResponse& responseOut) noexcept
{
    responseOut = {};

    const HINTERNET session = GetSharedWinHttpSession();
    if (! session)
    {
        return HRESULT_FROM_WIN32(ERROR_WINHTTP_INTERNAL_ERROR);
    }

    for (int attempt = 0; attempt < 4; ++attempt)
    {
        URL_COMPONENTS components{};
        std::wstring host;
        std::wstring pathAndQuery;
        HRESULT hr = CrackUrl(url, components, host, pathAndQuery);
        if (FAILED(hr))
        {
            Debug::Warning(L"Microsoft Drive: CrackUrl failed. method='{}' url='{}' hr=0x{:08X}", method, url, static_cast<unsigned long>(hr));
            return hr;
        }

        wil::unique_winhttp_hinternet connection(WinHttpConnect(session, host.c_str(), components.nPort, 0));
        if (! connection)
        {
            const DWORD lastError = GetLastError();
            Debug::Warning(L"Microsoft Drive: WinHttpConnect failed. method='{}' url='{}' lastError={}", method, url, lastError);
            return HRESULT_FROM_WIN32(lastError);
        }

        const DWORD requestFlags = (components.nScheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
        wil::unique_winhttp_hinternet request(WinHttpOpenRequest(
            connection.get(), std::wstring(method).c_str(), pathAndQuery.c_str(), nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, requestFlags));
        if (! request)
        {
            const DWORD lastError = GetLastError();
            Debug::Warning(L"Microsoft Drive: WinHttpOpenRequest failed. method='{}' url='{}' lastError={}", method, url, lastError);
            return HRESULT_FROM_WIN32(lastError);
        }

        const int timeoutMs = static_cast<int>(std::min<uint32_t>(settings.requestTimeoutMs, static_cast<uint32_t>((std::numeric_limits<int>::max)())));
        const int connectMs = static_cast<int>(std::min<uint32_t>(settings.connectTimeoutMs, static_cast<uint32_t>((std::numeric_limits<int>::max)())));
        static_cast<void>(WinHttpSetTimeouts(request.get(), connectMs, connectMs, timeoutMs, timeoutMs));

        if (bodySizeBytes > static_cast<size_t>((std::numeric_limits<DWORD>::max)()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        std::wstring headerBlock;
        for (const HttpHeader& header : headers)
        {
            headerBlock.append(header.name);
            headerBlock.append(L": ");
            headerBlock.append(header.value);
            headerBlock.append(L"\r\n");
        }

        std::wstring authorizationHeader;
        if (bearerToken && bearerToken[0] != '\0')
        {
            authorizationHeader = L"Bearer ";
            authorizationHeader.append(Utf16FromUtf8(bearerToken));
            headerBlock.append(L"Authorization: ");
            headerBlock.append(authorizationHeader);
            headerBlock.append(L"\r\n");
        }

        const DWORD totalBytes = static_cast<DWORD>(bodySizeBytes);
        void* optionalBody     = WINHTTP_NO_REQUEST_DATA;
        DWORD optionalLength   = 0;
        if (bodyBytes && totalBytes != 0)
        {
            optionalBody   = const_cast<std::byte*>(bodyBytes);
            optionalLength = totalBytes;
        }

        if (WinHttpSendRequest(request.get(),
                               headerBlock.empty() ? WINHTTP_NO_ADDITIONAL_HEADERS : headerBlock.c_str(),
                               headerBlock.empty() ? 0u : static_cast<DWORD>(headerBlock.size()),
                               optionalBody,
                               optionalLength,
                               totalBytes,
                               0) == 0)
        {
            const DWORD lastError = GetLastError();
            Debug::Warning(L"Microsoft Drive: WinHttpSendRequest failed. method='{}' url='{}' lastError={} bodyBytes={} headers='{}'",
                           method,
                           url,
                           lastError,
                           bodySizeBytes,
                           headerBlock.empty() ? std::wstring_view(L"<none>") : std::wstring_view(headerBlock));
            return HRESULT_FROM_WIN32(lastError);
        }

        if (bodyFile && bodySizeBytes != 0)
        {
            hr = ResetFilePointerToStart(bodyFile);
            if (FAILED(hr))
            {
                return hr;
            }

            std::array<std::byte, 64 * 1024> chunk{};
            uint64_t remaining = bodySizeBytes;
            while (remaining > 0)
            {
                const DWORD chunkSize = static_cast<DWORD>(std::min<uint64_t>(remaining, chunk.size()));
                DWORD read            = 0;
                if (ReadFile(bodyFile, chunk.data(), chunkSize, &read, nullptr) == 0)
                {
                    const DWORD lastError = GetLastError();
                    Debug::Warning(L"Microsoft Drive: ReadFile failed while sending request body. method='{}' url='{}' lastError={}", method, url, lastError);
                    return HRESULT_FROM_WIN32(lastError);
                }

                if (read == 0)
                {
                    return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
                }

                DWORD written = 0;
                if (WinHttpWriteData(request.get(), chunk.data(), read, &written) == 0)
                {
                    const DWORD lastError = GetLastError();
                    Debug::Warning(L"Microsoft Drive: WinHttpWriteData failed. method='{}' url='{}' lastError={} chunkBytes={}", method, url, lastError, read);
                    return HRESULT_FROM_WIN32(lastError);
                }

                if (written != read)
                {
                    return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
                }

                remaining -= static_cast<uint64_t>(read);
            }
        }

        if (WinHttpReceiveResponse(request.get(), nullptr) == 0)
        {
            const DWORD lastError = GetLastError();
            Debug::Warning(L"Microsoft Drive: WinHttpReceiveResponse failed. method='{}' url='{}' lastError={}", method, url, lastError);
            return HRESULT_FROM_WIN32(lastError);
        }

        DWORD statusCode = 0;
        DWORD statusSize = sizeof(statusCode);
        if (WinHttpQueryHeaders(request.get(),
                                WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                                WINHTTP_HEADER_NAME_BY_INDEX,
                                &statusCode,
                                &statusSize,
                                WINHTTP_NO_HEADER_INDEX) == 0)
        {
            const DWORD lastError = GetLastError();
            Debug::Warning(L"Microsoft Drive: WinHttpQueryHeaders(status) failed. method='{}' url='{}' lastError={}", method, url, lastError);
            return HRESULT_FROM_WIN32(lastError);
        }

        responseOut.statusCode  = statusCode;
        responseOut.retryAfter  = QueryStringHeader(request.get(), WINHTTP_QUERY_RETRY_AFTER);
        responseOut.requestId   = QueryStringHeader(request.get(), WINHTTP_QUERY_CUSTOM, L"request-id");
        responseOut.location    = QueryStringHeader(request.get(), WINHTTP_QUERY_LOCATION);
        responseOut.contentType = QueryStringHeader(request.get(), WINHTTP_QUERY_CONTENT_TYPE);
        responseOut.body.clear();

        while (true)
        {
            DWORD available = 0;
            if (WinHttpQueryDataAvailable(request.get(), &available) == 0)
            {
                const DWORD lastError = GetLastError();
                Debug::Warning(L"Microsoft Drive: WinHttpQueryDataAvailable failed. method='{}' url='{}' lastError={}", method, url, lastError);
                return HRESULT_FROM_WIN32(lastError);
            }

            if (available == 0)
            {
                break;
            }

            const size_t start = responseOut.body.size();
            responseOut.body.resize(start + available);
            DWORD read = 0;
            if (WinHttpReadData(request.get(), responseOut.body.data() + start, available, &read) == 0)
            {
                const DWORD lastError = GetLastError();
                Debug::Warning(L"Microsoft Drive: WinHttpReadData failed. method='{}' url='{}' lastError={}", method, url, lastError);
                return HRESULT_FROM_WIN32(lastError);
            }
            responseOut.body.resize(start + read);
        }

        if (allowRetry && (statusCode == 429u || statusCode == 503u || statusCode == 504u) && attempt < 3)
        {
            uint64_t delayMs = kDefaultThrottleDelayMs;
            if (! ParseRetryAfterMs(responseOut.retryAfter, delayMs))
            {
                delayMs = kDefaultThrottleDelayMs * static_cast<uint64_t>(attempt + 1);
            }

            Debug::Info(L"Microsoft Drive: retrying request after throttle. method='{}' url='{}' status={} delayMs={}", method, url, statusCode, delayMs);
            std::this_thread::sleep_for(std::chrono::milliseconds(delayMs));
            continue;
        }

        return S_OK;
    }

    return E_FAIL;
}

[[nodiscard]] HRESULT DownloadHttpToFile(
    const FileSystemMicrosoftDrive::Settings& settings, std::wstring_view url, const char* bearerToken, HANDLE file, HttpResponse& responseOut) noexcept
{
    responseOut = {};

    const HINTERNET session = GetSharedWinHttpSession();
    if (! session)
    {
        return HRESULT_FROM_WIN32(ERROR_WINHTTP_INTERNAL_ERROR);
    }

    URL_COMPONENTS components{};
    std::wstring host;
    std::wstring pathAndQuery;
    HRESULT hr = CrackUrl(url, components, host, pathAndQuery);
    if (FAILED(hr))
    {
        return hr;
    }

    wil::unique_winhttp_hinternet connection(WinHttpConnect(session, host.c_str(), components.nPort, 0));
    if (! connection)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const DWORD requestFlags = (components.nScheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
    wil::unique_winhttp_hinternet request(
        WinHttpOpenRequest(connection.get(), L"GET", pathAndQuery.c_str(), nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, requestFlags));
    if (! request)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    const int timeoutMs = static_cast<int>(std::min<uint32_t>(settings.requestTimeoutMs, static_cast<uint32_t>((std::numeric_limits<int>::max)())));
    const int connectMs = static_cast<int>(std::min<uint32_t>(settings.connectTimeoutMs, static_cast<uint32_t>((std::numeric_limits<int>::max)())));
    static_cast<void>(WinHttpSetTimeouts(request.get(), connectMs, connectMs, timeoutMs, timeoutMs));

    std::wstring headers;
    if (bearerToken && bearerToken[0] != '\0')
    {
        headers = L"Authorization: Bearer ";
        headers.append(Utf16FromUtf8(bearerToken));
        headers.append(L"\r\n");
    }

    if (WinHttpSendRequest(request.get(),
                           headers.empty() ? WINHTTP_NO_ADDITIONAL_HEADERS : headers.c_str(),
                           headers.empty() ? 0u : static_cast<DWORD>(headers.size()),
                           WINHTTP_NO_REQUEST_DATA,
                           0,
                           0,
                           0) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    if (WinHttpReceiveResponse(request.get(), nullptr) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    DWORD statusCode = 0;
    DWORD statusSize = sizeof(statusCode);
    if (WinHttpQueryHeaders(request.get(),
                            WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                            WINHTTP_HEADER_NAME_BY_INDEX,
                            &statusCode,
                            &statusSize,
                            WINHTTP_NO_HEADER_INDEX) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    responseOut.statusCode = statusCode;
    responseOut.requestId  = QueryStringHeader(request.get(), WINHTTP_QUERY_CUSTOM, L"request-id");
    responseOut.body.clear();

    std::array<std::byte, 64 * 1024> buffer{};
    while (true)
    {
        DWORD read = 0;
        if (WinHttpReadData(request.get(), buffer.data(), static_cast<DWORD>(buffer.size()), &read) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (read == 0)
        {
            break;
        }

        if (statusCode >= 200u && statusCode < 300u)
        {
            const HRESULT writeHr = WriteBufferToFile(file, buffer.data(), read);
            if (FAILED(writeHr))
            {
                return writeHr;
            }
        }
        else
        {
            responseOut.body.append(reinterpret_cast<const char*>(buffer.data()), read);
        }
    }

    return S_OK;
}

[[nodiscard]] HRESULT HresultFromGraphError(DWORD statusCode, std::string_view bodyUtf8) noexcept
{
    std::wstring code;
    if (! bodyUtf8.empty())
    {
        yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
        if (doc)
        {
            auto freeDoc     = wil::scope_exit([&] { yyjson_doc_free(doc); });
            yyjson_val* root = yyjson_doc_get_root(doc);
            if (root && yyjson_is_obj(root))
            {
                if (yyjson_val* error = yyjson_obj_get(root, "error"); error && yyjson_is_obj(error))
                {
                    code = TryGetJsonString(error, "code").value_or(L"");
                }
            }
        }
    }

    if (statusCode == 400u)
    {
        if (OrdinalString::EqualsNoCase(code, L"nameAlreadyExists"))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        return E_INVALIDARG;
    }
    if (statusCode == 401u)
    {
        return HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE);
    }
    if (statusCode == 403u)
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
    if (statusCode == 404u)
    {
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }
    if (statusCode == 409u)
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }
    if (statusCode == 412u)
    {
        return HRESULT_FROM_WIN32(ERROR_RETRY);
    }
    if (statusCode == 429u || statusCode == 503u || statusCode == 504u)
    {
        return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
    }
    if (statusCode == 507u || OrdinalString::EqualsNoCase(code, L"quotaLimitReached"))
    {
        return HRESULT_FROM_WIN32(ERROR_DISK_FULL);
    }

    return E_FAIL;
}

[[nodiscard]] std::wstring BuildTokenEndpoint(std::wstring_view authority) noexcept
{
    return std::format(L"https://{}/{}/oauth2/v2.0/token", kAuthHost, PercentEncodeUtf8(authority));
}

[[nodiscard]] std::wstring BuildAuthorizeUrl(std::wstring_view clientId,
                                             std::wstring_view authority,
                                             std::wstring_view redirectUri,
                                             std::wstring_view scopeText,
                                             std::wstring_view verifierChallenge,
                                             std::wstring_view state,
                                             std::wstring_view loginHint) noexcept
{
    std::wstring url = std::format(
        L"https://{}/{}/oauth2/v2.0/"
        L"authorize?client_id={}&response_type=code&redirect_uri={}&response_mode=query&scope={}&state={}&code_challenge={}&code_challenge_method=S256",
        kAuthHost,
        PercentEncodeUtf8(authority),
        PercentEncodeUtf8(clientId),
        PercentEncodeUtf8(redirectUri),
        PercentEncodeUtf8(scopeText),
        PercentEncodeUtf8(state),
        PercentEncodeUtf8(verifierChallenge));
    if (! loginHint.empty())
    {
        url.append(L"&login_hint=");
        url.append(PercentEncodeUtf8(loginHint));
    }

    return url;
}

[[nodiscard]] HRESULT ParseTokenResponse(std::string_view bodyUtf8, TokenResponse& tokenOut) noexcept
{
    tokenOut = {};

    yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
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

    if (yyjson_val* error = yyjson_obj_get(root, "error"); error && yyjson_is_str(error))
    {
        tokenOut.error            = TryGetJsonString(root, "error").value_or(L"");
        tokenOut.errorDescription = TryGetJsonString(root, "error_description").value_or(L"");
        return S_OK;
    }

    const auto accessToken = TryGetJsonString(root, "access_token");
    if (! accessToken.has_value())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    tokenOut.accessToken = Utf8FromUtf16(*accessToken);
    if (tokenOut.accessToken.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    tokenOut.refreshToken     = TryGetJsonString(root, "refresh_token").value_or(L"");
    tokenOut.expiresInSeconds = TryGetJsonUInt(root, "expires_in").value_or(0);
    return S_OK;
}

[[nodiscard]] HRESULT ExchangeTokenRequest(const FileSystemMicrosoftDrive::Settings& settings,
                                           std::wstring_view authority,
                                           std::wstring_view formBody,
                                           TokenResponse& tokenOut) noexcept
{
    const std::string formUtf8 = Utf8FromUtf16(formBody);
    if (formUtf8.empty() && ! formBody.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    const std::wstring url                  = BuildTokenEndpoint(authority);
    const std::array<HttpHeader, 2> headers = {
        HttpHeader{L"Content-Type", L"application/x-www-form-urlencoded"},
        HttpHeader{L"Accept", L"application/json"},
    };

    HttpResponse response{};
    const HRESULT hr = SendHttpRequest(
        settings, L"POST", url, nullptr, headers, reinterpret_cast<const std::byte*>(formUtf8.data()), formUtf8.size(), nullptr, true, response);
    if (FAILED(hr))
    {
        return hr;
    }

    const HRESULT parseHr = ParseTokenResponse(response.body, tokenOut);
    if (FAILED(parseHr))
    {
        return parseHr;
    }

    if (response.statusCode < 200u || response.statusCode >= 300u || ! tokenOut.error.empty())
    {
        const std::wstring_view oauthError = tokenOut.error.empty() ? std::wstring_view(L"<none>") : std::wstring_view(tokenOut.error);
        const std::wstring_view oauthErrorDescription =
            tokenOut.errorDescription.empty() ? std::wstring_view(L"<none>") : std::wstring_view(tokenOut.errorDescription);
        Debug::Warning(L"Microsoft Drive OAuth token request failed. authority='{}' status={} oauthError='{}' oauthErrorDescription='{}'",
                       authority,
                       response.statusCode,
                       oauthError,
                       oauthErrorDescription);
        return HRESULT_FROM_WIN32(OrdinalString::EqualsNoCase(tokenOut.error, L"invalid_grant") ? ERROR_LOGON_FAILURE : ERROR_ACCESS_DENIED);
    }

    return S_OK;
}

[[nodiscard]] HRESULT RefreshAccessToken(const FileSystemMicrosoftDrive::Settings& settings,
                                         std::wstring_view clientId,
                                         std::wstring_view authority,
                                         std::wstring_view scopeText,
                                         std::wstring_view refreshToken,
                                         TokenResponse& tokenOut) noexcept
{
    const std::wstring formBody = std::format(L"client_id={}&scope={}&refresh_token={}&grant_type=refresh_token",
                                              PercentEncodeUtf8(clientId),
                                              PercentEncodeUtf8(scopeText),
                                              PercentEncodeUtf8(refreshToken));
    return ExchangeTokenRequest(settings, authority, formBody, tokenOut);
}

[[nodiscard]] HRESULT ExchangeAuthorizationCode(const FileSystemMicrosoftDrive::Settings& settings,
                                                std::wstring_view clientId,
                                                std::wstring_view authority,
                                                std::wstring_view scopeText,
                                                std::wstring_view redirectUri,
                                                std::wstring_view authorizationCode,
                                                std::wstring_view codeVerifier,
                                                TokenResponse& tokenOut) noexcept
{
    const std::wstring formBody = std::format(L"client_id={}&scope={}&code={}&redirect_uri={}&grant_type=authorization_code&code_verifier={}",
                                              PercentEncodeUtf8(clientId),
                                              PercentEncodeUtf8(scopeText),
                                              PercentEncodeUtf8(authorizationCode),
                                              PercentEncodeUtf8(redirectUri),
                                              PercentEncodeUtf8(codeVerifier));
    return ExchangeTokenRequest(settings, authority, formBody, tokenOut);
}

[[nodiscard]] HRESULT LaunchInteractiveAuth(const FileSystemMicrosoftDrive::Settings& settings,
                                            std::wstring_view clientId,
                                            std::wstring_view authority,
                                            std::wstring_view scopeText,
                                            std::wstring_view loginHint,
                                            TokenResponse& tokenOut) noexcept
{
    std::wstring verifier;
    std::wstring challenge;
    std::wstring state;
    HRESULT hr = BuildPkceValues(verifier, challenge, state);
    if (FAILED(hr))
    {
        return hr;
    }
    auto clearSecrets = wil::scope_exit([&]
    {
        SecureClear(verifier);
        SecureClear(state);
    });

    std::wstring redirectUri;
    AuthCodeResult authResult{};
    HRESULT redirectWaitHr = S_OK;
    std::thread listenerThread([&]() noexcept
    {
        const HRESULT localHr = WaitForAuthRedirect(kLoopbackPath, state, redirectUri, authResult);
        if (FAILED(localHr))
        {
            redirectWaitHr   = localHr;
            authResult.error = L"redirect_wait_failed";
        }
    });

    while (redirectUri.empty() && authResult.error.empty())
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(25));
    }

    if (! redirectUri.empty())
    {
        Debug::Info(L"Microsoft Drive: launching interactive auth. clientId='{}' authority='{}' redirectUri='{}' scope='{}' loginHintPresent={}",
                    clientId,
                    authority,
                    redirectUri,
                    scopeText,
                    loginHint.empty() ? L"false" : L"true");
        const std::wstring authorizeUrl = BuildAuthorizeUrl(clientId, authority, redirectUri, scopeText, challenge, state, loginHint);
        HINSTANCE shellResult           = ShellExecuteW(nullptr, L"open", authorizeUrl.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
        if (reinterpret_cast<INT_PTR>(shellResult) <= 32)
        {
            Debug::Warning(L"Microsoft Drive: failed to open browser for interactive auth. shellResult={} clientId='{}' authority='{}' redirectUri='{}'",
                           static_cast<long long>(reinterpret_cast<INT_PTR>(shellResult)),
                           clientId,
                           authority,
                           redirectUri);
            authResult.error = L"shell_execute_failed";
        }
    }

    if (listenerThread.joinable())
    {
        listenerThread.join();
    }

    if (! authResult.error.empty())
    {
        const std::wstring_view authErrorDescription =
            authResult.errorDescription.empty() ? std::wstring_view(L"<none>") : std::wstring_view(authResult.errorDescription);
        Debug::Warning(L"Microsoft Drive: interactive auth failed. clientId='{}' authority='{}' redirectUri='{}' scope='{}' authError='{}' "
                       L"authErrorDescription='{}' listenerHr=0x{:08X}",
                       clientId,
                       authority,
                       redirectUri,
                       scopeText,
                       authResult.error,
                       authErrorDescription,
                       static_cast<unsigned long>(redirectWaitHr));
        return HRESULT_FROM_WIN32(OrdinalString::EqualsNoCase(authResult.error, L"access_denied") ? ERROR_CANCELLED : ERROR_LOGON_FAILURE);
    }

    if (redirectUri.empty() || authResult.code.empty() || ! OrdinalString::EqualsNoCase(authResult.state, state))
    {
        return HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE);
    }

    const HRESULT exchangeHr = ExchangeAuthorizationCode(settings, clientId, authority, scopeText, redirectUri, authResult.code, verifier, tokenOut);
    if (FAILED(exchangeHr))
    {
        Debug::Warning(L"Microsoft Drive: authorization-code exchange failed. clientId='{}' authority='{}' redirectUri='{}' scope='{}' hr=0x{:08X}",
                       clientId,
                       authority,
                       redirectUri,
                       scopeText,
                       static_cast<unsigned long>(exchangeHr));
    }

    return exchangeHr;
}

[[nodiscard]] HRESULT ResolveConnectionProfile(IHostConnections* hostConnections,
                                               FileSystemMicrosoftDriveMode mode,
                                               std::wstring_view connectionName,
                                               ConnectionProfileInfo& profileOut) noexcept
{
    profileOut = {};

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

    if (! json || json.get()[0] == '\0')
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

    profileOut.id           = TryGetJsonString(root, "id").value_or(L"");
    profileOut.name         = TryGetJsonString(root, "name").value_or(std::wstring(connectionName));
    profileOut.pluginId     = TryGetJsonString(root, "pluginId").value_or(L"");
    profileOut.host         = TryGetJsonString(root, "host").value_or(L"");
    profileOut.initialPath  = TryGetJsonString(root, "initialPath").value_or(L"/");
    profileOut.userName     = TryGetJsonString(root, "userName").value_or(L"");
    profileOut.authMode     = TryGetJsonString(root, "authMode").value_or(L"password");
    profileOut.savePassword = TryGetJsonBool(root, "savePassword").value_or(false);

    const wchar_t* expectedPluginId = nullptr;
    switch (mode)
    {
        case FileSystemMicrosoftDriveMode::OneDrivePersonal: expectedPluginId = L"builtin/file-system-onedrive-personal"; break;
        case FileSystemMicrosoftDriveMode::OneDriveBusiness: expectedPluginId = L"builtin/file-system-onedrive-business"; break;
        case FileSystemMicrosoftDriveMode::SharePoint: expectedPluginId = L"builtin/file-system-sharepoint"; break;
    }

    if (! expectedPluginId || ! OrdinalString::EqualsNoCase(profileOut.pluginId, expectedPluginId))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_NAME);
    }

    if (! OrdinalString::EqualsNoCase(profileOut.authMode, L"oauth2Pkce"))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if (yyjson_val* extra = yyjson_obj_get(root, "extra"); extra && yyjson_is_obj(extra))
    {
        profileOut.authorityOverride = TryGetJsonString(extra, "tenantAuthority").value_or(L"");
        profileOut.driveId           = TryGetJsonString(extra, "driveId").value_or(L"");
    }

    profileOut.initialPath = NormalizePluginPath(profileOut.initialPath.empty() ? std::wstring_view(L"/") : std::wstring_view(profileOut.initialPath));
    return S_OK;
}

[[nodiscard]] HRESULT ResolveConnectionFromPluginPath(FileSystemMicrosoftDriveMode mode,
                                                      IHostConnections* hostConnections,
                                                      std::wstring_view rawPath,
                                                      ConnectionProfileInfo& profileOut,
                                                      std::wstring& connectionNameOut,
                                                      std::wstring& drivePathOut) noexcept
{
    connectionNameOut.clear();
    drivePathOut.clear();

    const std::wstring canonicalPath = CanonicalizeInputPath(rawPath);
    if (canonicalPath == L"/" || canonicalPath.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_CONNECTED);
    }

    constexpr std::wstring_view kConnPrefix = L"/@conn:";
    if (! OrdinalString::StartsWithNoCase(canonicalPath, kConnPrefix))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_CONNECTED);
    }

    const std::wstring_view rest = std::wstring_view(canonicalPath).substr(kConnPrefix.size());
    const size_t nextSlash       = rest.find(L'/');
    const std::wstring_view name = nextSlash == std::wstring_view::npos ? rest : rest.substr(0, nextSlash);
    if (name.empty())
    {
        return E_INVALIDARG;
    }

    connectionNameOut.assign(name);
    HRESULT hr = ResolveConnectionProfile(hostConnections, mode, connectionNameOut, profileOut);
    if (FAILED(hr))
    {
        return hr;
    }

    if (nextSlash == std::wstring_view::npos)
    {
        drivePathOut = profileOut.initialPath;
    }
    else
    {
        drivePathOut = NormalizePluginPath(rest.substr(nextSlash));
    }

    if (drivePathOut.empty())
    {
        drivePathOut = L"/";
    }

    return S_OK;
}

[[nodiscard]] std::wstring DetermineAuthority(const ConnectionProfileInfo& profile, FileSystemMicrosoftDriveMode mode) noexcept
{
    if (! profile.authorityOverride.empty())
    {
        return profile.authorityOverride;
    }

    switch (mode)
    {
        case FileSystemMicrosoftDriveMode::OneDrivePersonal: return L"consumers";
        case FileSystemMicrosoftDriveMode::OneDriveBusiness: return L"organizations";
        case FileSystemMicrosoftDriveMode::SharePoint: return L"organizations";
        default: return L"organizations";
    }
}

[[nodiscard]] std::wstring DetermineScope(FileSystemMicrosoftDriveMode mode) noexcept
{
    switch (mode)
    {
        case FileSystemMicrosoftDriveMode::OneDrivePersonal: return std::wstring(kScopeOneDrivePersonal);
        case FileSystemMicrosoftDriveMode::OneDriveBusiness: return std::wstring(kScopeOneDriveBusiness);
        case FileSystemMicrosoftDriveMode::SharePoint: return std::wstring(kScopeSharePoint);
        default: return std::wstring(kScopeOneDriveBusiness);
    }
}

[[nodiscard]] HRESULT ParseSharePointHost(std::wstring_view rawHost, std::wstring& hostNameOut, std::wstring& sitePathOut) noexcept
{
    hostNameOut.clear();
    sitePathOut.clear();

    std::wstring host(rawHost);
    if (host.empty())
    {
        return E_INVALIDARG;
    }

    for (wchar_t& ch : host)
    {
        if (ch == L'\\')
        {
            ch = L'/';
        }
    }

    constexpr std::wstring_view kHttps = L"https://";
    constexpr std::wstring_view kHttp  = L"http://";
    if (OrdinalString::StartsWithNoCase(host, kHttps))
    {
        host.erase(0, kHttps.size());
    }
    else if (OrdinalString::StartsWithNoCase(host, kHttp))
    {
        host.erase(0, kHttp.size());
    }

    const size_t slash = host.find(L'/');
    if (slash == std::wstring::npos)
    {
        hostNameOut = host;
    }
    else
    {
        hostNameOut = host.substr(0, slash);
        sitePathOut = NormalizePluginPath(host.substr(slash));
        if (sitePathOut == L"/")
        {
            sitePathOut.clear();
        }
    }

    return hostNameOut.empty() ? E_INVALIDARG : S_OK;
}

[[nodiscard]] bool UsesMeDriveEndpoints(const DriveContext& context) noexcept
{
    return OrdinalString::EqualsNoCase(context.profile.pluginId, L"builtin/file-system-onedrive-personal") ||
           OrdinalString::EqualsNoCase(context.profile.pluginId, L"builtin/file-system-onedrive-business");
}

[[nodiscard]] std::wstring BuildGraphDriveBaseUrl(const DriveContext& context) noexcept
{
    if (UsesMeDriveEndpoints(context))
    {
        return std::format(L"{}/me/drive", kGraphBaseUrl);
    }

    return std::format(L"{}/drives/{}", kGraphBaseUrl, PercentEncodeUtf8(context.driveId));
}

[[nodiscard]] std::wstring BuildGraphItemMetadataUrl(const DriveContext& context, std::wstring_view drivePath, bool includeDownloadUrl) noexcept
{
    const std::wstring trimmedPath  = TrimTrailingSlashPreserveRoot(NormalizePluginPath(drivePath));
    const std::wstring driveBaseUrl = BuildGraphDriveBaseUrl(context);

    if (includeDownloadUrl)
    {
        if (trimmedPath == L"/" || trimmedPath.empty())
        {
            return std::format(L"{}/root", driveBaseUrl);
        }

        return std::format(L"{}/root:/{}:", driveBaseUrl, PercentEncodeGraphPath(trimmedPath));
    }

    const std::wstring select = L"$select=id,name,size,createdDateTime,lastModifiedDateTime,file,folder,webUrl";
    if (trimmedPath == L"/" || trimmedPath.empty())
    {
        return std::format(L"{}/root?{}", driveBaseUrl, select);
    }

    return std::format(L"{}/root:/{}:?{}", driveBaseUrl, PercentEncodeGraphPath(trimmedPath), select);
}

[[nodiscard]] std::wstring BuildGraphChildrenUrl(const DriveContext& context, std::wstring_view drivePath, uint32_t pageSize) noexcept
{
    const std::wstring trimmedPath  = TrimTrailingSlashPreserveRoot(NormalizePluginPath(drivePath));
    const std::wstring query        = std::format(L"$top={}&$select=id,name,size,createdDateTime,lastModifiedDateTime,file,folder,webUrl", pageSize);
    const std::wstring driveBaseUrl = BuildGraphDriveBaseUrl(context);

    if (trimmedPath == L"/" || trimmedPath.empty())
    {
        return std::format(L"{}/root/children?{}", driveBaseUrl, query);
    }

    return std::format(L"{}/root:/{}:/children?{}", driveBaseUrl, PercentEncodeGraphPath(trimmedPath), query);
}

[[nodiscard]] std::wstring BuildGraphCreateDirectoryUrl(const DriveContext& context, std::wstring_view parentItemId) noexcept
{
    if (UsesMeDriveEndpoints(context))
    {
        return std::format(L"{}/me/drive/items/{}/children", kGraphBaseUrl, PercentEncodeUtf8(parentItemId));
    }

    return std::format(L"{}/drives/{}/items/{}/children", kGraphBaseUrl, PercentEncodeUtf8(context.driveId), PercentEncodeUtf8(parentItemId));
}

[[nodiscard]] std::wstring BuildGraphItemByIdUrl(const DriveContext& context, std::wstring_view itemId) noexcept
{
    if (UsesMeDriveEndpoints(context))
    {
        return std::format(L"{}/me/drive/items/{}", kGraphBaseUrl, PercentEncodeUtf8(itemId));
    }

    return std::format(L"{}/drives/{}/items/{}", kGraphBaseUrl, PercentEncodeUtf8(context.driveId), PercentEncodeUtf8(itemId));
}

[[nodiscard]] std::wstring BuildGraphUploadContentUrl(const DriveContext& context, std::wstring_view drivePath) noexcept
{
    const std::wstring trimmedPath  = TrimTrailingSlashPreserveRoot(NormalizePluginPath(drivePath));
    const std::wstring driveBaseUrl = BuildGraphDriveBaseUrl(context);
    return std::format(L"{}/root:/{}:/content", driveBaseUrl, PercentEncodeGraphPath(trimmedPath));
}

[[nodiscard]] std::wstring BuildGraphCreateUploadSessionUrl(const DriveContext& context, std::wstring_view drivePath) noexcept
{
    const std::wstring trimmedPath  = TrimTrailingSlashPreserveRoot(NormalizePluginPath(drivePath));
    const std::wstring driveBaseUrl = BuildGraphDriveBaseUrl(context);
    return std::format(L"{}/root:/{}:/createUploadSession", driveBaseUrl, PercentEncodeGraphPath(trimmedPath));
}

[[nodiscard]] HRESULT ParseItemMetadata(std::string_view bodyUtf8, ItemMetadata& itemOut) noexcept
{
    itemOut = {};

    yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
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

    itemOut.id             = TryGetJsonString(root, "id").value_or(L"");
    itemOut.name           = TryGetJsonString(root, "name").value_or(L"");
    itemOut.sizeBytes      = TryGetJsonUInt(root, "size").value_or(0);
    itemOut.creationTime   = ParseIso8601FileTime(TryGetJsonString(root, "createdDateTime").value_or(L""));
    itemOut.lastWriteTime  = ParseIso8601FileTime(TryGetJsonString(root, "lastModifiedDateTime").value_or(L""));
    itemOut.lastAccessTime = itemOut.lastWriteTime;
    itemOut.changeTime     = itemOut.lastWriteTime;
    itemOut.downloadUrl    = TryGetJsonString(root, "@microsoft.graph.downloadUrl").value_or(L"");
    itemOut.attributes     = yyjson_obj_get(root, "folder") != nullptr ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;
    itemOut.isFolder       = (itemOut.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    itemOut.rawJson.assign(bodyUtf8.data(), bodyUtf8.size());
    return itemOut.id.empty() ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA) : S_OK;
}

[[nodiscard]] HRESULT ParseDriveInfo(std::string_view bodyUtf8, std::wstring& driveIdOut, std::wstring& displayNameOut, std::wstring& webUrlOut) noexcept
{
    driveIdOut.clear();
    displayNameOut.clear();
    webUrlOut.clear();

    yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
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

    driveIdOut     = TryGetJsonString(root, "id").value_or(L"");
    displayNameOut = TryGetJsonString(root, "name").value_or(L"");
    webUrlOut      = TryGetJsonString(root, "webUrl").value_or(L"");
    return driveIdOut.empty() ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA) : S_OK;
}

[[nodiscard]] HRESULT ParseSiteId(std::string_view bodyUtf8, std::wstring& siteIdOut, std::wstring& displayNameOut) noexcept
{
    siteIdOut.clear();
    displayNameOut.clear();

    yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
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

    siteIdOut      = TryGetJsonString(root, "id").value_or(L"");
    displayNameOut = TryGetJsonString(root, "displayName").value_or(L"");
    return siteIdOut.empty() ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA) : S_OK;
}

[[nodiscard]] HRESULT ParseChildren(std::string_view bodyUtf8,
                                    std::vector<FilesInformationMicrosoftDrive::Entry>& entriesOut,
                                    std::wstring& nextLinkOut) noexcept
{
    entriesOut.clear();
    nextLinkOut.clear();

    yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
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

    nextLinkOut       = TryGetJsonString(root, "@odata.nextLink").value_or(L"");
    yyjson_val* value = yyjson_obj_get(root, "value");
    if (! value || ! yyjson_is_arr(value))
    {
        return S_OK;
    }

    size_t index     = 0;
    size_t max       = 0;
    yyjson_val* item = nullptr;
    yyjson_arr_foreach(value, index, max, item)
    {
        if (! item || ! yyjson_is_obj(item))
        {
            continue;
        }

        FilesInformationMicrosoftDrive::Entry entry{};
        entry.name           = TryGetJsonString(item, "name").value_or(L"");
        entry.sizeBytes      = TryGetJsonUInt(item, "size").value_or(0);
        entry.creationTime   = ParseIso8601FileTime(TryGetJsonString(item, "createdDateTime").value_or(L""));
        entry.lastWriteTime  = ParseIso8601FileTime(TryGetJsonString(item, "lastModifiedDateTime").value_or(L""));
        entry.lastAccessTime = entry.lastWriteTime;
        entry.changeTime     = entry.lastWriteTime;
        entry.attributes     = yyjson_obj_get(item, "folder") != nullptr ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;
        entriesOut.push_back(std::move(entry));
    }

    return S_OK;
}

[[nodiscard]] HRESULT WriteMutableJsonDocument(yyjson_mut_doc* doc, std::string& jsonUtf8Out) noexcept
{
    if (! doc)
    {
        return E_INVALIDARG;
    }

    size_t length = 0;
    char* rawJson = yyjson_mut_write(doc, 0, &length);
    if (! rawJson)
    {
        return E_OUTOFMEMORY;
    }
    auto freeJson = wil::scope_exit([&] { free(rawJson); });

    jsonUtf8Out.assign(rawJson, length);
    return S_OK;
}

[[nodiscard]] HRESULT BuildJsonBodyForDirectory(std::wstring_view name, std::string& jsonUtf8Out) noexcept
{
    const std::string nameUtf8 = Utf8FromUtf16(name);
    if (nameUtf8.empty() && ! name.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }
    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    yyjson_mut_doc_set_root(doc, root);
    static_cast<void>(yyjson_mut_obj_add_strcpy(doc, root, "name", nameUtf8.c_str()));
    static_cast<void>(yyjson_mut_obj_add_strcpy(doc, root, "@microsoft.graph.conflictBehavior", "fail"));
    static_cast<void>(yyjson_mut_obj_add_val(doc, root, "folder", yyjson_mut_obj(doc)));
    return WriteMutableJsonDocument(doc, jsonUtf8Out);
}

[[nodiscard]] HRESULT BuildJsonBodyForMoveRename(std::wstring_view newName, std::wstring_view parentId, bool includeParent, std::string& jsonUtf8Out) noexcept
{
    const std::string nameUtf8   = Utf8FromUtf16(newName);
    const std::string parentUtf8 = Utf8FromUtf16(parentId);
    if ((nameUtf8.empty() && ! newName.empty()) || (includeParent && parentUtf8.empty()))
    {
        return HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
    }

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }
    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    yyjson_mut_doc_set_root(doc, root);
    static_cast<void>(yyjson_mut_obj_add_strcpy(doc, root, "name", nameUtf8.c_str()));
    if (includeParent)
    {
        yyjson_mut_val* parentReference = yyjson_mut_obj(doc);
        static_cast<void>(yyjson_mut_obj_add_strcpy(doc, parentReference, "id", parentUtf8.c_str()));
        static_cast<void>(yyjson_mut_obj_add_val(doc, root, "parentReference", parentReference));
    }

    return WriteMutableJsonDocument(doc, jsonUtf8Out);
}

[[nodiscard]] HRESULT BuildJsonBodyForUploadSession(bool allowOverwrite, std::string& jsonUtf8Out) noexcept
{
    jsonUtf8Out = allowOverwrite ? R"json({"item":{"@microsoft.graph.conflictBehavior":"replace"}})json"
                                 : R"json({"item":{"@microsoft.graph.conflictBehavior":"fail"}})json";
    return S_OK;
}

[[nodiscard]] HRESULT ParseUploadUrl(std::string_view bodyUtf8, std::wstring& uploadUrlOut) noexcept
{
    uploadUrlOut.clear();

    yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
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

    uploadUrlOut = TryGetJsonString(root, "uploadUrl").value_or(L"");
    return uploadUrlOut.empty() ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA) : S_OK;
}

[[nodiscard]] HRESULT BuildDriveContext(FileSystemMicrosoftDrive& fs, std::wstring_view rawPath, DriveContext& contextOut) noexcept
{
    contextOut = {};

    ConnectionProfileInfo profile;
    std::wstring connectionName;
    std::wstring drivePath;
    HRESULT hr = ResolveConnectionFromPluginPath(fs.Mode(), fs.GetHostConnections(), rawPath, profile, connectionName, drivePath);
    if (FAILED(hr))
    {
        return hr;
    }

    contextOut.profile             = std::move(profile);
    contextOut.connectionName      = std::move(connectionName);
    contextOut.drivePath           = std::move(drivePath);
    contextOut.authority           = DetermineAuthority(contextOut.profile, fs.Mode());
    contextOut.scopeText           = DetermineScope(fs.Mode());
    contextOut.persistRefreshToken = contextOut.profile.savePassword;

    if (fs.TryGetCachedDrive(
            contextOut.connectionName, contextOut.siteId, contextOut.driveId, contextOut.driveDisplayName, contextOut.driveVolumeLabel, contextOut.driveWebUrl))
    {
        fs.UpdateDriveInfoSnapshot(contextOut.driveDisplayName, contextOut.driveVolumeLabel);
        return S_OK;
    }

    std::string accessToken;
    hr = fs.AcquireAccessTokenForConnection(
        contextOut.connectionName, contextOut.profile.userName, contextOut.authority, contextOut.scopeText, contextOut.persistRefreshToken, accessToken);
    if (FAILED(hr))
    {
        return hr;
    }
    auto clearAccessToken = wil::scope_exit([&] { SecureClear(accessToken); });

    const std::array<HttpHeader, 1> headers = {HttpHeader{L"Accept", L"application/json"}};
    if (fs.Mode() == FileSystemMicrosoftDriveMode::OneDrivePersonal || fs.Mode() == FileSystemMicrosoftDriveMode::OneDriveBusiness)
    {
        HttpResponse response{};
        hr = SendHttpRequest(fs.SnapshotSettings(),
                             L"GET",
                             std::format(L"{}/me/drive?$select=id,name,webUrl", kGraphBaseUrl),
                             accessToken.c_str(),
                             headers,
                             nullptr,
                             0,
                             nullptr,
                             true,
                             response);
        if (FAILED(hr))
        {
            return hr;
        }
        if (response.statusCode < 200u || response.statusCode >= 300u)
        {
            return HresultFromGraphError(response.statusCode, response.body);
        }

        hr = ParseDriveInfo(response.body, contextOut.driveId, contextOut.driveDisplayName, contextOut.driveWebUrl);
        if (FAILED(hr))
        {
            return hr;
        }
        contextOut.driveVolumeLabel = contextOut.driveDisplayName;
    }
    else
    {
        std::wstring hostName;
        std::wstring sitePath;
        hr = ParseSharePointHost(contextOut.profile.host, hostName, sitePath);
        if (FAILED(hr))
        {
            return hr;
        }

        std::wstring siteUrl;
        if (sitePath.empty())
        {
            siteUrl = std::format(L"{}/sites/{}?$select=id,displayName", kGraphBaseUrl, PercentEncodeUtf8(hostName));
        }
        else
        {
            siteUrl = std::format(L"{}/sites/{}:{}?$select=id,displayName", kGraphBaseUrl, PercentEncodeUtf8(hostName), PercentEncodeGraphPath(sitePath));
        }

        HttpResponse siteResponse{};
        hr = SendHttpRequest(fs.SnapshotSettings(), L"GET", siteUrl, accessToken.c_str(), headers, nullptr, 0, nullptr, true, siteResponse);
        if (FAILED(hr))
        {
            return hr;
        }
        if (siteResponse.statusCode < 200u || siteResponse.statusCode >= 300u)
        {
            return HresultFromGraphError(siteResponse.statusCode, siteResponse.body);
        }

        std::wstring siteDisplayName;
        hr = ParseSiteId(siteResponse.body, contextOut.siteId, siteDisplayName);
        if (FAILED(hr))
        {
            return hr;
        }

        const std::wstring driveUrl = ! contextOut.profile.driveId.empty()
                                          ? std::format(L"{}/drives/{}?$select=id,name,webUrl", kGraphBaseUrl, PercentEncodeUtf8(contextOut.profile.driveId))
                                          : std::format(L"{}/sites/{}/drive?$select=id,name,webUrl", kGraphBaseUrl, PercentEncodeUtf8(contextOut.siteId));

        HttpResponse driveResponse{};
        hr = SendHttpRequest(fs.SnapshotSettings(), L"GET", driveUrl, accessToken.c_str(), headers, nullptr, 0, nullptr, true, driveResponse);
        if (FAILED(hr))
        {
            return hr;
        }
        if (driveResponse.statusCode < 200u || driveResponse.statusCode >= 300u)
        {
            return HresultFromGraphError(driveResponse.statusCode, driveResponse.body);
        }

        hr = ParseDriveInfo(driveResponse.body, contextOut.driveId, contextOut.driveDisplayName, contextOut.driveWebUrl);
        if (FAILED(hr))
        {
            return hr;
        }
        contextOut.driveVolumeLabel = siteDisplayName.empty() ? contextOut.driveDisplayName : siteDisplayName;
    }

    fs.UpdateDriveInfoSnapshot(contextOut.driveDisplayName, contextOut.driveVolumeLabel);
    fs.StoreCachedDrive(
        contextOut.connectionName, contextOut.siteId, contextOut.driveId, contextOut.driveDisplayName, contextOut.driveVolumeLabel, contextOut.driveWebUrl);

    return S_OK;
}

[[nodiscard]] HRESULT SendGraphJsonRequest(FileSystemMicrosoftDrive& fs,
                                           const DriveContext& context,
                                           std::wstring_view method,
                                           std::wstring_view url,
                                           std::string_view bodyUtf8,
                                           std::span<const HttpHeader> extraHeaders,
                                           bool allowRetry,
                                           HttpResponse& responseOut) noexcept
{
    for (int attempt = 0; attempt < 2; ++attempt)
    {
        std::string accessToken;
        HRESULT hr = fs.AcquireAccessTokenForConnection(
            context.connectionName, context.profile.userName, context.authority, context.scopeText, context.persistRefreshToken, accessToken);
        if (FAILED(hr))
        {
            return hr;
        }
        auto clearToken = wil::scope_exit([&] { SecureClear(accessToken); });

        std::vector<HttpHeader> headers;
        headers.reserve(extraHeaders.size() + 2u);
        headers.push_back(HttpHeader{L"Accept", L"application/json"});
        if (! bodyUtf8.empty())
        {
            headers.push_back(HttpHeader{L"Content-Type", L"application/json"});
        }
        for (const HttpHeader& header : extraHeaders)
        {
            headers.push_back(header);
        }

        hr = SendHttpRequest(fs.SnapshotSettings(),
                             method,
                             url,
                             accessToken.c_str(),
                             headers,
                             bodyUtf8.empty() ? nullptr : reinterpret_cast<const std::byte*>(bodyUtf8.data()),
                             bodyUtf8.size(),
                             nullptr,
                             allowRetry,
                             responseOut);
        if (FAILED(hr))
        {
            return hr;
        }

        if (responseOut.statusCode != 401u || attempt != 0)
        {
            return S_OK;
        }

        fs.ClearCachedAccessToken(context.connectionName);
    }

    return S_OK;
}

[[nodiscard]] HRESULT GetItemMetadata(
    FileSystemMicrosoftDrive& fs, const DriveContext& context, std::wstring_view drivePath, bool includeDownloadUrl, ItemMetadata& itemOut) noexcept
{
    const std::wstring url = BuildGraphItemMetadataUrl(context, drivePath, includeDownloadUrl);
    HttpResponse response{};
    const HRESULT hr = SendGraphJsonRequest(fs, context, L"GET", url, {}, {}, true, response);
    if (FAILED(hr))
    {
        return hr;
    }

    if (response.statusCode < 200u || response.statusCode >= 300u)
    {
        Debug::Warning(L"Microsoft Drive: item metadata request failed. connection='{}' drivePath='{}' status={} requestId='{}' url='{}'",
                       context.connectionName,
                       std::wstring(drivePath),
                       response.statusCode,
                       response.requestId.empty() ? std::wstring_view(L"<none>") : std::wstring_view(response.requestId),
                       url);
        return HresultFromGraphError(response.statusCode, response.body);
    }

    return ParseItemMetadata(response.body, itemOut);
}

[[nodiscard]] HRESULT ListDirectory(FileSystemMicrosoftDrive& fs,
                                    const DriveContext& context,
                                    std::wstring_view drivePath,
                                    std::vector<FilesInformationMicrosoftDrive::Entry>& entriesOut) noexcept
{
    entriesOut.clear();

    const FileSystemMicrosoftDrive::Settings settings = fs.SnapshotSettings();
    std::wstring nextUrl                              = BuildGraphChildrenUrl(context, drivePath, settings.pageSize);
    while (! nextUrl.empty())
    {
        HttpResponse response{};
        const HRESULT hr = SendGraphJsonRequest(fs, context, L"GET", nextUrl, {}, {}, true, response);
        if (FAILED(hr))
        {
            return hr;
        }

        if (response.statusCode < 200u || response.statusCode >= 300u)
        {
            Debug::Warning(L"Microsoft Drive: directory listing request failed. connection='{}' drivePath='{}' status={} requestId='{}' url='{}'",
                           context.connectionName,
                           std::wstring(drivePath),
                           response.statusCode,
                           response.requestId.empty() ? std::wstring_view(L"<none>") : std::wstring_view(response.requestId),
                           nextUrl);
            return HresultFromGraphError(response.statusCode, response.body);
        }

        std::vector<FilesInformationMicrosoftDrive::Entry> pageEntries;
        std::wstring nextLink;
        const HRESULT parseHr = ParseChildren(response.body, pageEntries, nextLink);
        if (FAILED(parseHr))
        {
            return parseHr;
        }

        entriesOut.insert(entriesOut.end(), std::make_move_iterator(pageEntries.begin()), std::make_move_iterator(pageEntries.end()));
        nextUrl = std::move(nextLink);
    }

    return S_OK;
}

[[nodiscard]] HRESULT EnsureParentMetadata(FileSystemMicrosoftDrive& fs,
                                           const DriveContext& context,
                                           std::wstring_view parentPath,
                                           ItemMetadata& parentOut) noexcept
{
    const HRESULT hr = GetItemMetadata(fs, context, parentPath, false, parentOut);
    if (FAILED(hr))
    {
        return hr;
    }

    return parentOut.isFolder ? S_OK : HRESULT_FROM_WIN32(ERROR_DIRECTORY);
}

[[nodiscard]] HRESULT CreateDirectoryItem(FileSystemMicrosoftDrive& fs, const DriveContext& context, std::wstring_view drivePath) noexcept
{
    std::wstring parentPath;
    std::wstring leafName;
    HRESULT hr = SplitParentAndLeaf(drivePath, parentPath, leafName);
    if (FAILED(hr))
    {
        return hr;
    }

    ItemMetadata parent{};
    hr = EnsureParentMetadata(fs, context, parentPath, parent);
    if (FAILED(hr))
    {
        return hr;
    }

    std::string bodyUtf8;
    hr = BuildJsonBodyForDirectory(leafName, bodyUtf8);
    if (FAILED(hr))
    {
        return hr;
    }

    HttpResponse response{};
    hr = SendGraphJsonRequest(fs, context, L"POST", BuildGraphCreateDirectoryUrl(context, parent.id), bodyUtf8, {}, false, response);
    if (FAILED(hr))
    {
        return hr;
    }

    if (response.statusCode == 409u)
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    return (response.statusCode >= 200u && response.statusCode < 300u) ? S_OK : HresultFromGraphError(response.statusCode, response.body);
}

[[nodiscard]] HRESULT DeleteItemByPath(FileSystemMicrosoftDrive& fs, const DriveContext& context, std::wstring_view drivePath) noexcept
{
    ItemMetadata item{};
    HRESULT hr = GetItemMetadata(fs, context, drivePath, false, item);
    if (FAILED(hr))
    {
        return hr;
    }

    HttpResponse response{};
    hr = SendGraphJsonRequest(fs, context, L"DELETE", BuildGraphItemByIdUrl(context, item.id), {}, {}, false, response);
    if (FAILED(hr))
    {
        return hr;
    }

    return response.statusCode == 204u ? S_OK : HresultFromGraphError(response.statusCode, response.body);
}

[[nodiscard]] HRESULT MoveOrRenameItem(FileSystemMicrosoftDrive& fs,
                                       const DriveContext& sourceContext,
                                       const DriveContext& destinationContext,
                                       std::wstring_view sourcePath,
                                       std::wstring_view destinationPath,
                                       FileSystemFlags flags) noexcept
{
    if (! OrdinalString::EqualsNoCase(sourceContext.connectionName, destinationContext.connectionName) ||
        ! OrdinalString::EqualsNoCase(sourceContext.driveId, destinationContext.driveId))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    ItemMetadata sourceItem{};
    HRESULT hr = GetItemMetadata(fs, sourceContext, sourcePath, false, sourceItem);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring normalizedDestination = TrimTrailingSlashPreserveRoot(NormalizePluginPath(destinationPath));
    std::wstring destinationParentPath;
    std::wstring destinationName;
    hr = SplitParentAndLeaf(normalizedDestination, destinationParentPath, destinationName);
    if (FAILED(hr))
    {
        return hr;
    }

    if ((flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) == 0)
    {
        ItemMetadata existing{};
        if (SUCCEEDED(GetItemMetadata(fs, destinationContext, normalizedDestination, false, existing)))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
    }
    else
    {
        static_cast<void>(DeleteItemByPath(fs, destinationContext, normalizedDestination));
    }

    ItemMetadata destinationParent{};
    hr = EnsureParentMetadata(fs, destinationContext, destinationParentPath, destinationParent);
    if (FAILED(hr))
    {
        return hr;
    }

    const bool includeParent = ! OrdinalString::EqualsNoCase(destinationParent.id, sourceItem.id);
    std::string bodyUtf8;
    hr = BuildJsonBodyForMoveRename(destinationName, destinationParent.id, includeParent, bodyUtf8);
    if (FAILED(hr))
    {
        return hr;
    }

    HttpResponse response{};
    hr = SendGraphJsonRequest(fs, sourceContext, L"PATCH", BuildGraphItemByIdUrl(sourceContext, sourceItem.id), bodyUtf8, {}, false, response);
    if (FAILED(hr))
    {
        return hr;
    }

    return (response.statusCode >= 200u && response.statusCode < 300u) ? S_OK : HresultFromGraphError(response.statusCode, response.body);
}

[[nodiscard]] uint64_t ComputeUploadChunkSizeBytes(const FileSystemMicrosoftDrive::Settings& settings) noexcept
{
    uint64_t chunkBytes = static_cast<uint64_t>(settings.uploadChunkMiB) * 1024ull * 1024ull;
    chunkBytes          = std::max<uint64_t>(chunkBytes, kGraphChunkAlignmentBytes);
    chunkBytes          = (chunkBytes / kGraphChunkAlignmentBytes) * kGraphChunkAlignmentBytes;
    return chunkBytes == 0 ? kGraphChunkAlignmentBytes : chunkBytes;
}

[[nodiscard]] bool IsNotFoundStatus(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
}

[[nodiscard]] HRESULT CheckShouldCancel(IFileSystemCallback* callback, void* cookie, bool& cancelledOut) noexcept
{
    cancelledOut = false;
    if (! callback)
    {
        return S_OK;
    }

    BOOL cancel      = FALSE;
    const HRESULT hr = callback->FileSystemShouldCancel(&cancel, cookie);
    if (FAILED(hr))
    {
        return hr;
    }

    cancelledOut = cancel == TRUE;
    return S_OK;
}

[[nodiscard]] HRESULT ReportItemResult(IFileSystemCallback* callback,
                                       FileSystemOperation operationType,
                                       unsigned long totalItems,
                                       unsigned long completedItems,
                                       unsigned long itemIndex,
                                       const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       HRESULT status,
                                       const FileSystemOptions* options,
                                       void* cookie) noexcept
{
    if (! callback)
    {
        return S_OK;
    }

    auto* mutableOptions = const_cast<FileSystemOptions*>(options);
    HRESULT hr = callback->FileSystemProgress(operationType, totalItems, completedItems, 0, 0, sourcePath, destinationPath, 0, 0, mutableOptions, 0, cookie);
    if (FAILED(hr))
    {
        return hr;
    }

    return callback->FileSystemItemCompleted(operationType, itemIndex, sourcePath, destinationPath, status, mutableOptions, cookie);
}

class MicrosoftDriveRangedFileReader final : public IFileReader
{
public:
    MicrosoftDriveRangedFileReader(
        FileSystemMicrosoftDrive& fileSystem, DriveContext context, std::wstring drivePath, uint64_t sizeBytes, std::wstring downloadUrl) noexcept
        : _fileSystem(&fileSystem),
          _context(std::move(context)),
          _drivePath(std::move(drivePath)),
          _sizeBytes(sizeBytes),
          _downloadUrl(std::move(downloadUrl)),
          _settings(fileSystem.SnapshotSettings())
    {
        _fileSystem->AddRef();
    }

    MicrosoftDriveRangedFileReader(const MicrosoftDriveRangedFileReader&)            = delete;
    MicrosoftDriveRangedFileReader(MicrosoftDriveRangedFileReader&&)                 = delete;
    MicrosoftDriveRangedFileReader& operator=(const MicrosoftDriveRangedFileReader&) = delete;
    MicrosoftDriveRangedFileReader& operator=(MicrosoftDriveRangedFileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileReader))
        {
            *ppvObject = static_cast<IFileReader*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (result == 0)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (! sizeBytes)
        {
            return E_POINTER;
        }

        *sizeBytes = _sizeBytes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (! newPosition)
        {
            return E_POINTER;
        }

        *newPosition = 0;

        uint64_t base = 0;
        switch (origin)
        {
            case FILE_BEGIN: base = 0; break;
            case FILE_CURRENT: base = _position; break;
            case FILE_END: base = _sizeBytes; break;
            default: return E_INVALIDARG;
        }

        const __int64 signedBase = static_cast<__int64>(std::min<uint64_t>(base, static_cast<uint64_t>((std::numeric_limits<__int64>::max)())));
        const __int64 next       = signedBase + offset;
        if (next < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }

        _position    = static_cast<uint64_t>(next);
        *newPosition = _position;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (! bytesRead)
        {
            return E_POINTER;
        }

        *bytesRead = 0;
        if (bytesToRead == 0)
        {
            return S_OK;
        }
        if (! buffer)
        {
            return E_POINTER;
        }

        auto* output            = static_cast<std::byte*>(buffer);
        unsigned long totalRead = 0;
        while (totalRead < bytesToRead && _position < _sizeBytes)
        {
            const uint64_t cacheEnd = _cacheOffset + static_cast<uint64_t>(_cache.size());
            if (_cache.empty() || _position < _cacheOffset || _position >= cacheEnd)
            {
                const HRESULT hr = FetchRange(_position, bytesToRead - totalRead);
                if (FAILED(hr))
                {
                    return hr;
                }

                if (_cache.empty())
                {
                    break;
                }
            }

            const size_t cacheIndex = static_cast<size_t>(_position - _cacheOffset);
            if (cacheIndex >= _cache.size())
            {
                return HRESULT_FROM_WIN32(ERROR_READ_FAULT);
            }

            const size_t available = _cache.size() - cacheIndex;
            const size_t toCopy    = std::min<size_t>(available, bytesToRead - totalRead);
            memcpy(output + totalRead, _cache.data() + cacheIndex, toCopy);
            totalRead += static_cast<unsigned long>(toCopy);
            _position += static_cast<uint64_t>(toCopy);
        }

        *bytesRead = totalRead;
        return S_OK;
    }

private:
    ~MicrosoftDriveRangedFileReader()
    {
        if (_fileSystem)
        {
            _fileSystem->Release();
        }
    }

    [[nodiscard]] HRESULT RefreshDownloadUrl() noexcept
    {
        ItemMetadata item{};
        const HRESULT hr = GetItemMetadata(*_fileSystem, _context, _drivePath, true, item);
        if (FAILED(hr))
        {
            Debug::Warning(L"Microsoft Drive: failed to refresh download URL. connection='{}' drivePath='{}' hr=0x{:08X}",
                           _context.connectionName,
                           _drivePath,
                           static_cast<unsigned long>(hr));
            return hr;
        }

        if (item.isFolder)
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
        }

        _sizeBytes   = item.sizeBytes;
        _downloadUrl = std::move(item.downloadUrl);
        if (_downloadUrl.empty())
        {
            Debug::Warning(L"Microsoft Drive: metadata response did not contain download URL. connection='{}' drivePath='{}' itemId='{}'",
                           _context.connectionName,
                           _drivePath,
                           item.id.empty() ? std::wstring_view(L"<none>") : std::wstring_view(item.id));
        }
        return _downloadUrl.empty() ? HRESULT_FROM_WIN32(ERROR_INVALID_DATA) : S_OK;
    }

    [[nodiscard]] HRESULT FetchRange(uint64_t offset, unsigned long minimumBytes) noexcept
    {
        _cache.clear();

        if (offset >= _sizeBytes)
        {
            _cacheOffset = offset;
            return S_OK;
        }

        uint64_t chunkBytes = std::max<uint64_t>(minimumBytes, 256ull * 1024ull);
        chunkBytes          = std::min<uint64_t>(chunkBytes, 1024ull * 1024ull);
        chunkBytes          = std::min<uint64_t>(chunkBytes, _sizeBytes - offset);

        for (int attempt = 0; attempt < 2; ++attempt)
        {
            HRESULT hr = _downloadUrl.empty() ? RefreshDownloadUrl() : S_OK;
            if (FAILED(hr))
            {
                return hr;
            }

            const uint64_t endOffset                = offset + chunkBytes - 1ull;
            const std::array<HttpHeader, 2> headers = {
                HttpHeader{L"Accept", L"application/octet-stream"},
                HttpHeader{L"Range", std::format(L"bytes={}-{}", offset, endOffset)},
            };

            HttpResponse response{};
            hr = SendHttpRequest(_settings, L"GET", _downloadUrl, nullptr, headers, nullptr, 0, nullptr, true, response);
            if (FAILED(hr))
            {
                Debug::Warning(L"Microsoft Drive: ranged download request failed. connection='{}' drivePath='{}' offset={} bytes={} hr=0x{:08X}",
                               _context.connectionName,
                               _drivePath,
                               offset,
                               chunkBytes,
                               static_cast<unsigned long>(hr));
                return hr;
            }

            if ((response.statusCode == 401u || response.statusCode == 403u) && attempt == 0)
            {
                _downloadUrl.clear();
                continue;
            }

            if (response.statusCode != 200u && response.statusCode != 206u)
            {
                Debug::Warning(
                    L"Microsoft Drive: ranged download returned unexpected status. connection='{}' drivePath='{}' offset={} bytes={} status={} requestId='{}'",
                    _context.connectionName,
                    _drivePath,
                    offset,
                    chunkBytes,
                    response.statusCode,
                    response.requestId.empty() ? std::wstring_view(L"<none>") : std::wstring_view(response.requestId));
                return HresultFromGraphError(response.statusCode, response.body);
            }

            if (response.statusCode == 200u && offset != 0)
            {
                if (response.body.size() <= offset)
                {
                    return HRESULT_FROM_WIN32(ERROR_READ_FAULT);
                }
                _cacheOffset = 0;
            }
            else
            {
                _cacheOffset = offset;
            }
            _cache.resize(response.body.size());
            if (! response.body.empty())
            {
                memcpy(_cache.data(), response.body.data(), response.body.size());
            }
            return S_OK;
        }

        return HRESULT_FROM_WIN32(ERROR_READ_FAULT);
    }

    std::atomic_ulong _refCount{1};
    FileSystemMicrosoftDrive* _fileSystem = nullptr;
    DriveContext _context;
    std::wstring _drivePath;
    uint64_t _position    = 0;
    uint64_t _sizeBytes   = 0;
    uint64_t _cacheOffset = 0;
    std::wstring _downloadUrl;
    FileSystemMicrosoftDrive::Settings _settings;
    std::vector<std::byte> _cache;
};

class MicrosoftDriveFileWriter final : public IFileWriter
{
public:
    MicrosoftDriveFileWriter(FileSystemMicrosoftDrive& fileSystem, std::wstring destinationPath, FileSystemFlags flags, wil::unique_hfile tempFile) noexcept
        : _fileSystem(&fileSystem),
          _destinationPath(std::move(destinationPath)),
          _flags(flags),
          _tempFile(std::move(tempFile))
    {
        _fileSystem->AddRef();
    }

    MicrosoftDriveFileWriter(const MicrosoftDriveFileWriter&)            = delete;
    MicrosoftDriveFileWriter(MicrosoftDriveFileWriter&&)                 = delete;
    MicrosoftDriveFileWriter& operator=(const MicrosoftDriveFileWriter&) = delete;
    MicrosoftDriveFileWriter& operator=(MicrosoftDriveFileWriter&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (! ppvObject)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileWriter))
        {
            *ppvObject = static_cast<IFileWriter*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (result == 0)
        {
            delete this;
        }
        return result;
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
        if (! bytesWritten)
        {
            return E_POINTER;
        }

        *bytesWritten = 0;

        if (_committed)
        {
            return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
        }
        if (bytesToWrite == 0)
        {
            return S_OK;
        }
        if (! buffer)
        {
            return E_POINTER;
        }
        if (! _tempFile)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        DWORD written = 0;
        if (WriteFile(_tempFile.get(), buffer, bytesToWrite, &written, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        _position += static_cast<uint64_t>(written);
        *bytesWritten = written;
        return written == bytesToWrite ? S_OK : HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
    }

    HRESULT STDMETHODCALLTYPE Commit() noexcept override;

private:
    ~MicrosoftDriveFileWriter()
    {
        if (_fileSystem)
        {
            _fileSystem->Release();
        }
    }

    [[nodiscard]] HRESULT UploadSimple(const DriveContext& context, std::wstring_view drivePath, uint64_t fileSize, bool allowOverwrite) noexcept;
    [[nodiscard]] HRESULT UploadWithSession(const DriveContext& context, std::wstring_view drivePath, uint64_t fileSize, bool allowOverwrite) noexcept;

    std::atomic_ulong _refCount{1};
    FileSystemMicrosoftDrive* _fileSystem = nullptr;
    std::wstring _destinationPath;
    FileSystemFlags _flags = FILESYSTEM_FLAG_NONE;
    wil::unique_hfile _tempFile;
    uint64_t _position = 0;
    bool _committed    = false;
};

} // namespace FsMs

using namespace FsMs;

FileSystemMicrosoftDrive::FileSystemMicrosoftDrive(FileSystemMicrosoftDriveMode mode, IHost* host) : _mode(mode)
{
    switch (_mode)
    {
        case FileSystemMicrosoftDriveMode::OneDrivePersonal:
            _metaData.id          = kPluginIdOneDrivePersonal;
            _metaData.shortId     = kPluginShortIdOneDrivePersonal;
            _metaData.name        = LocalizedPluginName(FileSystemMicrosoftDriveMode::OneDrivePersonal);
            _metaData.description = LocalizedPluginDescription(FileSystemMicrosoftDriveMode::OneDrivePersonal);
            break;
        case FileSystemMicrosoftDriveMode::OneDriveBusiness:
            _metaData.id          = kPluginIdOneDriveBusiness;
            _metaData.shortId     = kPluginShortIdOneDriveBusiness;
            _metaData.name        = LocalizedPluginName(FileSystemMicrosoftDriveMode::OneDriveBusiness);
            _metaData.description = LocalizedPluginDescription(FileSystemMicrosoftDriveMode::OneDriveBusiness);
            break;
        case FileSystemMicrosoftDriveMode::SharePoint:
            _metaData.id          = kPluginIdSharePoint;
            _metaData.shortId     = kPluginShortIdSharePoint;
            _metaData.name        = LocalizedPluginName(FileSystemMicrosoftDriveMode::SharePoint);
            _metaData.description = LocalizedPluginDescription(FileSystemMicrosoftDriveMode::SharePoint);
            break;
    }

    _metaData.author  = kPluginAuthor;
    _metaData.version = kPluginVersion;

    _configurationJson = "{}";
    _propertiesJson    = "{}";
    _driveFileSystem   = _metaData.shortId ? _metaData.shortId : L"";

    if (host)
    {
        static_cast<void>(host->QueryInterface(__uuidof(IHostAlerts), _hostAlerts.put_void()));
        static_cast<void>(host->QueryInterface(__uuidof(IHostConnections), _hostConnections.put_void()));
    }
}

FileSystemMicrosoftDrive::~FileSystemMicrosoftDrive()
{
    std::lock_guard lock(_stateMutex);
    for (auto& [key, token] : _tokenCacheByConnectionName)
    {
        SecureClear(token.accessToken);
    }
    _tokenCacheByConnectionName.clear();
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::QueryInterface(REFIID riid, void** ppvObject) noexcept
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

ULONG STDMETHODCALLTYPE FileSystemMicrosoftDrive::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE FileSystemMicrosoftDrive::Release() noexcept
{
    const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (result == 0)
    {
        delete this;
    }
    return result;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (! metaData)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (! schemaJsonUtf8)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = kSchemaJson;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    std::lock_guard lock(_stateMutex);

    for (auto& [key, token] : _tokenCacheByConnectionName)
    {
        SecureClear(token.accessToken);
    }
    _tokenCacheByConnectionName.clear();
    _driveCacheByConnectionName.clear();
    _settings = {};

    if (! configurationJsonUtf8 || configurationJsonUtf8[0] == '\0')
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

    if (const auto value = TryGetJsonString(root, "clientId"); value.has_value())
    {
        _settings.clientId = value.value();
    }

    if (const auto value = TryGetJsonUInt(root, "connectTimeoutMs"); value.has_value() && value.value() >= 1u)
    {
        _settings.connectTimeoutMs = static_cast<uint32_t>(std::min<uint64_t>(value.value(), 600'000ull));
    }

    if (const auto value = TryGetJsonUInt(root, "requestTimeoutMs"); value.has_value() && value.value() >= 1u)
    {
        _settings.requestTimeoutMs = static_cast<uint32_t>(std::min<uint64_t>(value.value(), 600'000ull));
    }

    if (const auto value = TryGetJsonUInt(root, "pageSize"); value.has_value() && value.value() >= 1u)
    {
        _settings.pageSize = static_cast<uint32_t>(std::min<uint64_t>(value.value(), 999ull));
    }

    if (const auto value = TryGetJsonUInt(root, "uploadChunkMiB"); value.has_value() && value.value() >= 1u)
    {
        _settings.uploadChunkMiB = static_cast<uint32_t>(std::min<uint64_t>(value.value(), 32ull));
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (! configurationJsonUtf8)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (! pSomethingToSave)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    *pSomethingToSave = (! _configurationJson.empty() && _configurationJson != "{}") ? TRUE : FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetMenuItems(const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (! items || ! count)
    {
        return E_POINTER;
    }

    static NavigationMenuItem menuItems[3]{};
    menuItems[0].flags     = NAV_MENU_ITEM_FLAG_HEADER;
    menuItems[0].label     = _metaData.name;
    menuItems[0].path      = nullptr;
    menuItems[0].iconPath  = nullptr;
    menuItems[0].commandId = 0;

    menuItems[1].flags     = NAV_MENU_ITEM_FLAG_SEPARATOR;
    menuItems[1].label     = nullptr;
    menuItems[1].path      = nullptr;
    menuItems[1].iconPath  = nullptr;
    menuItems[1].commandId = 0;

    menuItems[2].flags     = NAV_MENU_ITEM_FLAG_NONE;
    menuItems[2].label     = L"/";
    menuItems[2].path      = L"/";
    menuItems[2].iconPath  = nullptr;
    menuItems[2].commandId = 0;

    *items = menuItems;
    *count = static_cast<unsigned int>(std::size(menuItems));
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::ExecuteMenuCommand([[maybe_unused]] unsigned int commandId) noexcept
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::SetCallback(INavigationMenuCallback* callback, void* cookie) noexcept
{
    std::lock_guard lock(_stateMutex);
    _navigationMenuCallback = callback;
    _navigationMenuCookie   = callback ? cookie : nullptr;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetDriveInfo(const wchar_t* path, DriveInfo* info) noexcept
{
    if (! info)
    {
        return E_POINTER;
    }

    const wchar_t* scheme    = _metaData.shortId ? _metaData.shortId : L"";
    std::wstring displayName = std::format(L"{}:/", scheme);
    std::wstring volumeLabel;

    if (path && path[0] != L'\0')
    {
        ConnectionProfileInfo profile{};
        std::wstring connectionName;
        std::wstring drivePath;
        if (SUCCEEDED(ResolveConnectionFromPluginPath(_mode, GetHostConnections(), path, profile, connectionName, drivePath)))
        {
            std::wstring siteId;
            std::wstring driveId;
            std::wstring cachedDisplayName;
            std::wstring cachedVolumeLabel;
            std::wstring cachedWebUrl;
            if (TryGetCachedDrive(connectionName, siteId, driveId, cachedDisplayName, cachedVolumeLabel, cachedWebUrl) && ! cachedDisplayName.empty())
            {
                displayName = cachedDisplayName;
                volumeLabel = cachedVolumeLabel;
            }
            else
            {
                displayName = std::format(L"{}://{}", scheme, connectionName);
            }

            if (! drivePath.empty() && drivePath != L"/")
            {
                displayName.append(drivePath);
            }
        }
    }

    {
        std::lock_guard lock(_stateMutex);
        _driveDisplayName = std::move(displayName);
        _driveVolumeLabel = std::move(volumeLabel);

        info->flags = static_cast<DriveInfoFlags>(DRIVE_INFO_FLAG_HAS_DISPLAY_NAME | DRIVE_INFO_FLAG_HAS_FILE_SYSTEM);
        if (! _driveVolumeLabel.empty())
        {
            info->flags = static_cast<DriveInfoFlags>(info->flags | DRIVE_INFO_FLAG_HAS_VOLUME_LABEL);
        }

        info->displayName = _driveDisplayName.empty() ? nullptr : _driveDisplayName.c_str();
        info->volumeLabel = _driveVolumeLabel.empty() ? nullptr : _driveVolumeLabel.c_str();
        info->fileSystem  = _driveFileSystem.empty() ? nullptr : _driveFileSystem.c_str();
        info->totalBytes  = 0;
        info->freeBytes   = 0;
        info->usedBytes   = 0;
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetDriveMenuItems([[maybe_unused]] const wchar_t* path,
                                                                      const NavigationMenuItem** items,
                                                                      unsigned int* count) noexcept
{
    if (! items || ! count)
    {
        return E_POINTER;
    }

    *items = nullptr;
    *count = 0;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::ExecuteDriveMenuCommand([[maybe_unused]] unsigned int commandId,
                                                                            [[maybe_unused]] const wchar_t* path) noexcept
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept
{
    if (! ppFilesInformation)
    {
        return E_POINTER;
    }

    *ppFilesInformation = nullptr;
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    DriveContext context{};
    HRESULT hr = BuildDriveContext(*this, path, context);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring normalizedDrivePath = TrimTrailingSlashPreserveRoot(NormalizePluginPath(context.drivePath));
    if (normalizedDrivePath != L"/")
    {
        ItemMetadata item{};
        hr = GetItemMetadata(*this, context, normalizedDrivePath, false, item);
        if (FAILED(hr))
        {
            return hr;
        }

        if (! item.isFolder)
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
        }
    }

    std::vector<FilesInformationMicrosoftDrive::Entry> entries;
    hr = ListDirectory(*this, context, normalizedDrivePath, entries);
    if (FAILED(hr))
    {
        return hr;
    }

    auto* info = new (std::nothrow) FilesInformationMicrosoftDrive();
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

    UpdateDriveInfoSnapshot(context.driveDisplayName, context.driveVolumeLabel);
    *ppFilesInformation = info;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::CopyItem([[maybe_unused]] const wchar_t* sourcePath,
                                                             [[maybe_unused]] const wchar_t* destinationPath,
                                                             [[maybe_unused]] FileSystemFlags flags,
                                                             [[maybe_unused]] const FileSystemOptions* options,
                                                             [[maybe_unused]] IFileSystemCallback* callback,
                                                             [[maybe_unused]] void* cookie) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::MoveItem(const wchar_t* sourcePath,
                                                             const wchar_t* destinationPath,
                                                             FileSystemFlags flags,
                                                             const FileSystemOptions* options,
                                                             IFileSystemCallback* callback,
                                                             void* cookie) noexcept
{
    if (! sourcePath || ! destinationPath || sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    DriveContext sourceContext{};
    HRESULT hr = BuildDriveContext(*this, sourcePath, sourceContext);
    if (FAILED(hr))
    {
        return hr;
    }

    DriveContext destinationContext{};
    hr = BuildDriveContext(*this, destinationPath, destinationContext);
    if (FAILED(hr))
    {
        return hr;
    }

    hr                       = MoveOrRenameItem(*this, sourceContext, destinationContext, sourceContext.drivePath, destinationContext.drivePath, flags);
    const HRESULT callbackHr = ReportItemResult(callback, FILESYSTEM_MOVE, 1, 1, 0, sourcePath, destinationPath, hr, options, cookie);
    return FAILED(callbackHr) ? callbackHr : hr;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::DeleteItem(
    const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options, IFileSystemCallback* callback, void* cookie) noexcept
{
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    DriveContext context{};
    HRESULT hr = BuildDriveContext(*this, path, context);
    if (FAILED(hr))
    {
        return hr;
    }

    if ((flags & FILESYSTEM_FLAG_RECURSIVE) == 0)
    {
        ItemMetadata item{};
        hr = GetItemMetadata(*this, context, context.drivePath, false, item);
        if (FAILED(hr))
        {
            return hr;
        }

        if (item.isFolder)
        {
            std::vector<FilesInformationMicrosoftDrive::Entry> entries;
            hr = ListDirectory(*this, context, context.drivePath, entries);
            if (FAILED(hr))
            {
                return hr;
            }
            if (! entries.empty())
            {
                hr                       = HRESULT_FROM_WIN32(ERROR_DIR_NOT_EMPTY);
                const HRESULT callbackHr = ReportItemResult(callback, FILESYSTEM_DELETE, 1, 1, 0, path, nullptr, hr, options, cookie);
                return FAILED(callbackHr) ? callbackHr : hr;
            }
        }
    }

    hr                       = DeleteItemByPath(*this, context, context.drivePath);
    const HRESULT callbackHr = ReportItemResult(callback, FILESYSTEM_DELETE, 1, 1, 0, path, nullptr, hr, options, cookie);
    return FAILED(callbackHr) ? callbackHr : hr;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::RenameItem(const wchar_t* sourcePath,
                                                               const wchar_t* destinationPath,
                                                               FileSystemFlags flags,
                                                               const FileSystemOptions* options,
                                                               IFileSystemCallback* callback,
                                                               void* cookie) noexcept
{
    if (! sourcePath || ! destinationPath || sourcePath[0] == L'\0' || destinationPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    DriveContext sourceContext{};
    HRESULT hr = BuildDriveContext(*this, sourcePath, sourceContext);
    if (FAILED(hr))
    {
        return hr;
    }

    DriveContext destinationContext{};
    hr = BuildDriveContext(*this, destinationPath, destinationContext);
    if (FAILED(hr))
    {
        return hr;
    }

    hr                       = MoveOrRenameItem(*this, sourceContext, destinationContext, sourceContext.drivePath, destinationContext.drivePath, flags);
    const HRESULT callbackHr = ReportItemResult(callback, FILESYSTEM_RENAME, 1, 1, 0, sourcePath, destinationPath, hr, options, cookie);
    return FAILED(callbackHr) ? callbackHr : hr;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::CopyItems(const wchar_t* const* sourcePaths,
                                                              unsigned long count,
                                                              const wchar_t* destinationFolder,
                                                              FileSystemFlags flags,
                                                              const FileSystemOptions* options,
                                                              IFileSystemCallback* callback,
                                                              void* cookie) noexcept
{
    if ((! sourcePaths && count != 0) || ! destinationFolder || destinationFolder[0] == L'\0')
    {
        return (! sourcePaths && count != 0) ? E_POINTER : E_INVALIDARG;
    }

    const std::wstring destinationRoot = CanonicalizeInputPath(destinationFolder);
    HRESULT firstFailure               = S_OK;
    for (unsigned long i = 0; i < count; ++i)
    {
        bool cancelled = false;
        HRESULT hr     = CheckShouldCancel(callback, cookie, cancelled);
        if (FAILED(hr))
        {
            return hr;
        }
        if (cancelled)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        std::wstring parentPath;
        std::wstring leafName;
        hr = SplitParentAndLeaf(CanonicalizeInputPath(sourcePaths[i]), parentPath, leafName);
        if (SUCCEEDED(hr))
        {
            const std::wstring destinationPath = JoinPath(destinationRoot, leafName);
            hr                                 = CopyItem(sourcePaths[i], destinationPath.c_str(), flags, options, nullptr, nullptr);
            const HRESULT callbackHr =
                ReportItemResult(callback, FILESYSTEM_COPY, count, i + 1u, i, sourcePaths[i], destinationPath.c_str(), hr, options, cookie);
            if (FAILED(callbackHr))
            {
                return callbackHr;
            }
        }
        else
        {
            const HRESULT callbackHr = ReportItemResult(callback, FILESYSTEM_COPY, count, i + 1u, i, sourcePaths[i], destinationFolder, hr, options, cookie);
            if (FAILED(callbackHr))
            {
                return callbackHr;
            }
        }
        if (FAILED(hr) && SUCCEEDED(firstFailure))
        {
            firstFailure = hr;
            if ((flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) == 0)
            {
                return hr;
            }
        }
    }

    return firstFailure;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::MoveItems(const wchar_t* const* sourcePaths,
                                                              unsigned long count,
                                                              const wchar_t* destinationFolder,
                                                              FileSystemFlags flags,
                                                              const FileSystemOptions* options,
                                                              IFileSystemCallback* callback,
                                                              void* cookie) noexcept
{
    if ((! sourcePaths && count != 0) || ! destinationFolder || destinationFolder[0] == L'\0')
    {
        return (! sourcePaths && count != 0) ? E_POINTER : E_INVALIDARG;
    }

    const std::wstring destinationRoot = CanonicalizeInputPath(destinationFolder);
    HRESULT firstFailure               = S_OK;
    for (unsigned long i = 0; i < count; ++i)
    {
        bool cancelled = false;
        HRESULT hr     = CheckShouldCancel(callback, cookie, cancelled);
        if (FAILED(hr))
        {
            return hr;
        }
        if (cancelled)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        std::wstring parentPath;
        std::wstring leafName;
        hr = SplitParentAndLeaf(CanonicalizeInputPath(sourcePaths[i]), parentPath, leafName);
        if (SUCCEEDED(hr))
        {
            const std::wstring destinationPath = JoinPath(destinationRoot, leafName);
            hr                                 = MoveItem(sourcePaths[i], destinationPath.c_str(), flags, options, nullptr, nullptr);
            const HRESULT callbackHr =
                ReportItemResult(callback, FILESYSTEM_MOVE, count, i + 1u, i, sourcePaths[i], destinationPath.c_str(), hr, options, cookie);
            if (FAILED(callbackHr))
            {
                return callbackHr;
            }
        }
        else
        {
            const HRESULT callbackHr = ReportItemResult(callback, FILESYSTEM_MOVE, count, i + 1u, i, sourcePaths[i], destinationFolder, hr, options, cookie);
            if (FAILED(callbackHr))
            {
                return callbackHr;
            }
        }
        if (FAILED(hr) && SUCCEEDED(firstFailure))
        {
            firstFailure = hr;
            if ((flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) == 0)
            {
                return hr;
            }
        }
    }

    return firstFailure;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::DeleteItems(const wchar_t* const* paths,
                                                                unsigned long count,
                                                                FileSystemFlags flags,
                                                                const FileSystemOptions* options,
                                                                IFileSystemCallback* callback,
                                                                void* cookie) noexcept
{
    if (! paths && count != 0)
    {
        return E_POINTER;
    }

    HRESULT firstFailure = S_OK;
    for (unsigned long i = 0; i < count; ++i)
    {
        bool cancelled = false;
        HRESULT hr     = CheckShouldCancel(callback, cookie, cancelled);
        if (FAILED(hr))
        {
            return hr;
        }
        if (cancelled)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        hr                       = DeleteItem(paths[i], flags, options, nullptr, nullptr);
        const HRESULT callbackHr = ReportItemResult(callback, FILESYSTEM_DELETE, count, i + 1u, i, paths[i], nullptr, hr, options, cookie);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }
        if (FAILED(hr) && SUCCEEDED(firstFailure))
        {
            firstFailure = hr;
            if ((flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) == 0)
            {
                return hr;
            }
        }
    }

    return firstFailure;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::RenameItems(const FileSystemRenamePair* items,
                                                                unsigned long count,
                                                                FileSystemFlags flags,
                                                                const FileSystemOptions* options,
                                                                IFileSystemCallback* callback,
                                                                void* cookie) noexcept
{
    if (! items && count != 0)
    {
        return E_POINTER;
    }

    HRESULT firstFailure = S_OK;
    for (unsigned long i = 0; i < count; ++i)
    {
        bool cancelled = false;
        HRESULT hr     = CheckShouldCancel(callback, cookie, cancelled);
        if (FAILED(hr))
        {
            return hr;
        }
        if (cancelled)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        std::wstring parentPath;
        std::wstring existingLeafName;
        hr = SplitParentAndLeaf(CanonicalizeInputPath(items[i].sourcePath), parentPath, existingLeafName);
        std::wstring destinationPath;
        if (SUCCEEDED(hr))
        {
            if (! items[i].newName || items[i].newName[0] == L'\0' || wcschr(items[i].newName, L'/') || wcschr(items[i].newName, L'\\'))
            {
                hr = E_INVALIDARG;
            }
            else
            {
                destinationPath = JoinPath(parentPath, items[i].newName);
                hr              = RenameItem(items[i].sourcePath, destinationPath.c_str(), flags, options, nullptr, nullptr);
            }
        }

        const HRESULT callbackHr = ReportItemResult(callback,
                                                    FILESYSTEM_RENAME,
                                                    count,
                                                    i + 1u,
                                                    i,
                                                    items[i].sourcePath,
                                                    destinationPath.empty() ? nullptr : destinationPath.c_str(),
                                                    hr,
                                                    options,
                                                    cookie);
        if (FAILED(callbackHr))
        {
            return callbackHr;
        }
        if (FAILED(hr) && SUCCEEDED(firstFailure))
        {
            firstFailure = hr;
            if ((flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) == 0)
            {
                return hr;
            }
        }
    }

    return firstFailure;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetCapabilities(const char** jsonUtf8) noexcept
{
    if (! jsonUtf8)
    {
        return E_POINTER;
    }

    *jsonUtf8 = kCapabilitiesJson;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept
{
    if (! fileAttributes)
    {
        return E_POINTER;
    }

    *fileAttributes = 0;
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (CanonicalizeInputPath(path) == L"/")
    {
        *fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
        return S_OK;
    }

    DriveContext context{};
    const HRESULT hr = BuildDriveContext(*this, path, context);
    if (FAILED(hr))
    {
        return hr;
    }

    ItemMetadata item{};
    const HRESULT metadataHr = GetItemMetadata(*this, context, context.drivePath, false, item);
    if (FAILED(metadataHr))
    {
        return metadataHr;
    }

    *fileAttributes = item.attributes;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept
{
    if (! reader)
    {
        return E_POINTER;
    }

    *reader = nullptr;
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    DriveContext context{};
    HRESULT hr = BuildDriveContext(*this, path, context);
    if (FAILED(hr))
    {
        return hr;
    }

    ItemMetadata item{};
    hr = GetItemMetadata(*this, context, context.drivePath, true, item);
    if (FAILED(hr))
    {
        Debug::Warning(L"Microsoft Drive: CreateFileReader metadata lookup failed. connection='{}' drivePath='{}' hr=0x{:08X}",
                       context.connectionName,
                       context.drivePath,
                       static_cast<unsigned long>(hr));
        return hr;
    }
    if (item.isFolder)
    {
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
    }
    if (item.downloadUrl.empty())
    {
        Debug::Warning(L"Microsoft Drive: CreateFileReader missing download URL. connection='{}' drivePath='{}' itemId='{}'",
                       context.connectionName,
                       context.drivePath,
                       item.id.empty() ? std::wstring_view(L"<none>") : std::wstring_view(item.id));
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    const std::wstring drivePath = context.drivePath;
    auto* impl = new (std::nothrow) MicrosoftDriveRangedFileReader(*this, std::move(context), drivePath, item.sizeBytes, std::move(item.downloadUrl));
    if (! impl)
    {
        return E_OUTOFMEMORY;
    }

    *reader = impl;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept
{
    if (! writer)
    {
        return E_POINTER;
    }

    *writer = nullptr;
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const std::wstring normalizedPath = TrimTrailingSlashPreserveRoot(CanonicalizeInputPath(path));
    if (normalizedPath == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
    }

    wil::unique_hfile tempFile;
    const HRESULT tempHr = GetTemporaryDeleteOnCloseFile(tempFile);
    if (FAILED(tempHr))
    {
        return tempHr;
    }

    auto* impl = new (std::nothrow) MicrosoftDriveFileWriter(*this, normalizedPath, flags, std::move(tempFile));
    if (! impl)
    {
        return E_OUTOFMEMORY;
    }

    *writer = impl;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept
{
    if (! info)
    {
        return E_POINTER;
    }
    if (info->sizeBytes != sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    DriveContext context{};
    HRESULT hr = BuildDriveContext(*this, path, context);
    if (FAILED(hr))
    {
        return hr;
    }

    ItemMetadata item{};
    hr = GetItemMetadata(*this, context, context.drivePath, false, item);
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

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::SetFileBasicInformation([[maybe_unused]] const wchar_t* path,
                                                                            [[maybe_unused]] const FileSystemBasicInformation* info) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept
{
    if (! jsonUtf8)
    {
        return E_POINTER;
    }

    *jsonUtf8 = nullptr;
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    DriveContext context{};
    HRESULT hr = BuildDriveContext(*this, path, context);
    if (FAILED(hr))
    {
        return hr;
    }

    ItemMetadata item{};
    hr = GetItemMetadata(*this, context, context.drivePath, true, item);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = SetPropertiesJson(std::move(item.rawJson));
    if (FAILED(hr))
    {
        return hr;
    }

    std::lock_guard lock(_stateMutex);
    *jsonUtf8 = _propertiesJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::CreateDirectory(const wchar_t* path) noexcept
{
    if (! path || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    DriveContext context{};
    HRESULT hr = BuildDriveContext(*this, path, context);
    if (FAILED(hr))
    {
        return hr;
    }

    if (context.drivePath == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    return CreateDirectoryItem(*this, context, context.drivePath);
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetDirectorySize(
    const wchar_t* path, FileSystemFlags flags, IFileSystemDirectorySizeCallback* callback, void* cookie, FileSystemDirectorySizeResult* result) noexcept
{
    if (! result)
    {
        return E_POINTER;
    }
    if (result->sizeBytes != sizeof(FileSystemDirectorySizeResult))
    {
        return E_INVALIDARG;
    }

    result->totalBytes     = 0;
    result->fileCount      = 0;
    result->directoryCount = 0;
    result->status         = S_OK;

    if (! path || path[0] == L'\0')
    {
        result->status = E_INVALIDARG;
        return result->status;
    }

    DriveContext context{};
    HRESULT hr = BuildDriveContext(*this, path, context);
    if (FAILED(hr))
    {
        result->status = hr;
        return hr;
    }

    ItemMetadata rootItem{};
    hr = GetItemMetadata(*this, context, context.drivePath, false, rootItem);
    if (FAILED(hr))
    {
        result->status = hr;
        return hr;
    }

    if (! rootItem.isFolder)
    {
        result->totalBytes = rootItem.sizeBytes;
        result->fileCount  = 1;
        if (callback)
        {
            hr = callback->DirectorySizeProgress(1, result->totalBytes, result->fileCount, result->directoryCount, path, cookie);
            if (FAILED(hr))
            {
                result->status = hr;
            }
        }
        return result->status;
    }

    const bool recursive = (flags & FILESYSTEM_FLAG_RECURSIVE) != 0;
    std::vector<std::wstring> pending;
    pending.push_back(context.drivePath);
    uint64_t scannedEntries = 0;

    while (! pending.empty())
    {
        BOOL cancel = FALSE;
        if (callback)
        {
            hr = callback->DirectorySizeShouldCancel(&cancel, cookie);
            if (FAILED(hr))
            {
                result->status = hr;
                return hr;
            }
            if (cancel)
            {
                result->status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                return result->status;
            }
        }

        const std::wstring current = std::move(pending.back());
        pending.pop_back();

        std::vector<FilesInformationMicrosoftDrive::Entry> entries;
        hr = ListDirectory(*this, context, current, entries);
        if (FAILED(hr))
        {
            result->status = hr;
            return hr;
        }

        for (const auto& entry : entries)
        {
            ++scannedEntries;

            if ((entry.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0)
            {
                ++result->directoryCount;
                if (recursive)
                {
                    pending.push_back(JoinPath(current, entry.name));
                }
            }
            else
            {
                result->totalBytes += entry.sizeBytes;
                ++result->fileCount;
            }

            if (callback && (scannedEntries % 128u) == 0u)
            {
                const std::wstring currentPath = JoinPath(current, entry.name);
                hr =
                    callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, currentPath.c_str(), cookie);
                if (FAILED(hr))
                {
                    result->status = hr;
                    return hr;
                }
            }
        }

        if (! recursive)
        {
            break;
        }
    }

    if (callback)
    {
        hr = callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, nullptr, cookie);
        if (FAILED(hr))
        {
            result->status = hr;
            return hr;
        }
    }

    return S_OK;
}

HRESULT MicrosoftDriveFileWriter::UploadSimple(const DriveContext& context, std::wstring_view drivePath, uint64_t fileSize, bool allowOverwrite) noexcept
{
    std::string accessToken;
    HRESULT hr = _fileSystem->AcquireAccessTokenForConnection(
        context.connectionName, context.profile.userName, context.authority, context.scopeText, context.persistRefreshToken, accessToken);
    if (FAILED(hr))
    {
        return hr;
    }
    auto clearToken = wil::scope_exit([&] { SecureClear(accessToken); });

    std::vector<HttpHeader> headers;
    headers.push_back(HttpHeader{L"Accept", L"application/json"});
    headers.push_back(HttpHeader{L"Content-Type", L"application/octet-stream"});
    if (! allowOverwrite)
    {
        headers.push_back(HttpHeader{L"If-None-Match", L"*"});
    }

    HttpResponse response{};
    hr = SendHttpRequest(_fileSystem->SnapshotSettings(),
                         L"PUT",
                         BuildGraphUploadContentUrl(context, drivePath),
                         accessToken.c_str(),
                         headers,
                         nullptr,
                         static_cast<size_t>(fileSize),
                         _tempFile.get(),
                         true,
                         response);
    if (FAILED(hr))
    {
        return hr;
    }

    return (response.statusCode >= 200u && response.statusCode < 300u) ? S_OK : HresultFromGraphError(response.statusCode, response.body);
}

HRESULT MicrosoftDriveFileWriter::UploadWithSession(const DriveContext& context, std::wstring_view drivePath, uint64_t fileSize, bool allowOverwrite) noexcept
{
    std::string bodyUtf8;
    HRESULT hr = BuildJsonBodyForUploadSession(allowOverwrite, bodyUtf8);
    if (FAILED(hr))
    {
        return hr;
    }

    HttpResponse sessionResponse{};
    hr = SendGraphJsonRequest(*_fileSystem, context, L"POST", BuildGraphCreateUploadSessionUrl(context, drivePath), bodyUtf8, {}, false, sessionResponse);
    if (FAILED(hr))
    {
        return hr;
    }
    if (sessionResponse.statusCode < 200u || sessionResponse.statusCode >= 300u)
    {
        return HresultFromGraphError(sessionResponse.statusCode, sessionResponse.body);
    }

    std::wstring uploadUrl;
    hr = ParseUploadUrl(sessionResponse.body, uploadUrl);
    if (FAILED(hr))
    {
        return hr;
    }

    const FileSystemMicrosoftDrive::Settings settings = _fileSystem->SnapshotSettings();
    const uint64_t chunkBytes                         = ComputeUploadChunkSizeBytes(settings);
    if (chunkBytes > static_cast<uint64_t>((std::numeric_limits<unsigned long>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    hr = ResetFilePointerToStart(_tempFile.get());
    if (FAILED(hr))
    {
        return hr;
    }

    DWORD lastStatus = 0;
    std::vector<std::byte> buffer(static_cast<size_t>(chunkBytes));
    uint64_t offset = 0;
    while (offset < fileSize)
    {
        const uint64_t remaining = fileSize - offset;
        const DWORD toRead       = static_cast<DWORD>(std::min<uint64_t>(remaining, chunkBytes));

        DWORD read = 0;
        if (ReadFile(_tempFile.get(), buffer.data(), toRead, &read, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (read != toRead)
        {
            return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        }

        const uint64_t endOffset                = offset + static_cast<uint64_t>(read) - 1ull;
        const std::array<HttpHeader, 3> headers = {
            HttpHeader{L"Accept", L"application/json"},
            HttpHeader{L"Content-Length", std::format(L"{}", read)},
            HttpHeader{L"Content-Range", std::format(L"bytes {}-{}/{}", offset, endOffset, fileSize)},
        };

        HttpResponse uploadResponse{};
        hr = SendHttpRequest(settings, L"PUT", uploadUrl, nullptr, headers, buffer.data(), read, nullptr, true, uploadResponse);
        if (FAILED(hr))
        {
            return hr;
        }

        lastStatus = uploadResponse.statusCode;
        if (lastStatus != 200u && lastStatus != 201u && lastStatus != 202u)
        {
            return HresultFromGraphError(lastStatus, uploadResponse.body);
        }

        offset += static_cast<uint64_t>(read);
    }

    return (lastStatus == 200u || lastStatus == 201u) ? S_OK : HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
}

HRESULT MicrosoftDriveFileWriter::Commit() noexcept
{
    if (_committed)
    {
        return S_OK;
    }
    if (! _tempFile)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
    }

    DriveContext context{};
    HRESULT hr = BuildDriveContext(*_fileSystem, _destinationPath, context);
    if (FAILED(hr))
    {
        return hr;
    }

    const std::wstring normalizedPath = TrimTrailingSlashPreserveRoot(NormalizePluginPath(context.drivePath));
    if (normalizedPath == L"/")
    {
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
    }

    const bool allowOverwrite = (_flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
    if (! allowOverwrite)
    {
        ItemMetadata existing{};
        hr = GetItemMetadata(*_fileSystem, context, normalizedPath, false, existing);
        if (SUCCEEDED(hr))
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        if (! IsNotFoundStatus(hr))
        {
            return hr;
        }
    }

    uint64_t fileSize = 0;
    hr                = GetFileSizeBytes(_tempFile.get(), fileSize);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = (fileSize <= kGraphSimpleUploadMaxBytes) ? UploadSimple(context, normalizedPath, fileSize, allowOverwrite)
                                                  : UploadWithSession(context, normalizedPath, fileSize, allowOverwrite);
    if (SUCCEEDED(hr))
    {
        _committed = true;
    }
    return hr;
}

FileSystemMicrosoftDrive::Settings FileSystemMicrosoftDrive::SnapshotSettings() const noexcept
{
    std::lock_guard lock(_stateMutex);
    return _settings;
}

IHostConnections* FileSystemMicrosoftDrive::GetHostConnections() const noexcept
{
    return _hostConnections.get();
}

IHostAlerts* FileSystemMicrosoftDrive::GetHostAlerts() const noexcept
{
    return _hostAlerts.get();
}

void FileSystemMicrosoftDrive::ShowMissingClientIdAlert() const noexcept
{
    IHostAlerts* const hostAlerts = GetHostAlerts();
    if (! hostAlerts)
    {
        return;
    }

    std::wstring pluginName = _metaData.name ? _metaData.name : LocalizedPluginName(_mode);
    if (pluginName.empty())
    {
        pluginName = L"Microsoft Drive";
    }

    const std::wstring title   = LoadStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_ALERT_TITLE_SIGNIN_CONFIG_REQUIRED);
    const std::wstring message = FormatStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_ALERT_MSG_MISSING_CLIENT_ID_FMT, pluginName);
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

HRESULT FileSystemMicrosoftDrive::AcquireAccessTokenForConnection(std::wstring_view connectionName,
                                                                  std::wstring_view userName,
                                                                  std::wstring_view authority,
                                                                  std::wstring_view scopeText,
                                                                  bool persistRefreshToken,
                                                                  std::string& accessTokenOut) noexcept
{
    accessTokenOut.clear();

    const Settings settings = SnapshotSettings();
    if (settings.clientId.empty())
    {
        ShowMissingClientIdAlert();
        return HRESULT_FROM_WIN32(ERROR_BAD_CONFIGURATION);
    }

    const std::wstring cacheKey(connectionName);
    const uint64_t nowTickMs = GetTickCount64();

    {
        std::lock_guard lock(_stateMutex);
        const auto it = _tokenCacheByConnectionName.find(cacheKey);
        if (it != _tokenCacheByConnectionName.end() && ! it->second.accessToken.empty() && it->second.expiresAtTickMs > (nowTickMs + 30'000ull))
        {
            accessTokenOut = it->second.accessToken;
            return S_OK;
        }
    }

    std::wstring refreshToken;
    auto clearRefreshToken = wil::scope_exit([&] { SecureClear(refreshToken); });

    IHostConnections* hostConnections = GetHostConnections();
    if (hostConnections)
    {
        wchar_t* rawSecret     = nullptr;
        const HRESULT secretHr = hostConnections->GetConnectionSecret(cacheKey.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, nullptr, &rawSecret);
        if (SUCCEEDED(secretHr) && rawSecret)
        {
            const size_t rawLength = wcslen(rawSecret);
            refreshToken.assign(rawSecret, rawLength);
            SecureZeroMemory(rawSecret, rawLength * sizeof(wchar_t));
            CoTaskMemFree(rawSecret);
        }
        else if (FAILED(secretHr) && secretHr != HRESULT_FROM_WIN32(ERROR_NOT_FOUND))
        {
            return secretHr;
        }
    }

    auto cacheToken = [&](TokenResponse& token) noexcept -> HRESULT
    {
        if (token.accessToken.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE);
        }

        const uint64_t ttlSeconds    = std::max<uint64_t>(token.expiresInSeconds, 60ull);
        const uint64_t expiresAtTick = nowTickMs + (ttlSeconds * 1000ull);

        {
            std::lock_guard lock(_stateMutex);
            CachedToken& cached = _tokenCacheByConnectionName[cacheKey];
            SecureClear(cached.accessToken);
            cached.accessToken     = token.accessToken;
            cached.expiresAtTickMs = expiresAtTick;
        }

        accessTokenOut = token.accessToken;
        return S_OK;
    };

    if (! refreshToken.empty())
    {
        TokenResponse refreshed{};
        HRESULT hr = RefreshAccessToken(settings, settings.clientId, authority, scopeText, refreshToken, refreshed);
        if (SUCCEEDED(hr))
        {
            if (hostConnections && ! refreshed.refreshToken.empty())
            {
                const HRESULT storeHr = hostConnections->SetConnectionSecret(
                    cacheKey.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, refreshed.refreshToken.c_str(), persistRefreshToken ? TRUE : FALSE);
                if (FAILED(storeHr))
                {
                    Debug::Warning(
                        L"Microsoft Drive: Failed to store refreshed token for '{}' (hr=0x{:08X}).", cacheKey.c_str(), static_cast<unsigned long>(storeHr));
                }
            }

            return cacheToken(refreshed);
        }

        const std::wstring_view oauthError = refreshed.error.empty() ? std::wstring_view(L"<none>") : std::wstring_view(refreshed.error);
        const std::wstring_view oauthErrorDescription =
            refreshed.errorDescription.empty() ? std::wstring_view(L"<none>") : std::wstring_view(refreshed.errorDescription);
        Debug::Warning(L"Microsoft Drive: refresh-token auth failed for '{}' (hr=0x{:08X}) authority='{}' clientId='{}' scope='{}' oauthError='{}' "
                       L"oauthErrorDescription='{}'; falling back to interactive auth.",
                       cacheKey.c_str(),
                       static_cast<unsigned long>(hr),
                       authority,
                       settings.clientId,
                       scopeText,
                       oauthError,
                       oauthErrorDescription);

        if (hr == HRESULT_FROM_WIN32(ERROR_LOGON_FAILURE) && hostConnections)
        {
            static_cast<void>(hostConnections->DeleteConnectionSecret(cacheKey.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, TRUE));
        }
    }

    TokenResponse interactive{};
    Debug::Info(L"Microsoft Drive: starting interactive auth for '{}' authority='{}' clientId='{}' scope='{}' loginHintPresent={} persistRefreshToken={}",
                cacheKey.c_str(),
                authority,
                settings.clientId,
                scopeText,
                userName.empty() ? L"false" : L"true",
                persistRefreshToken ? L"true" : L"false");
    const HRESULT authHr = LaunchInteractiveAuth(settings, settings.clientId, authority, scopeText, userName, interactive);
    if (FAILED(authHr))
    {
        return authHr;
    }

    if (hostConnections && ! interactive.refreshToken.empty())
    {
        const HRESULT storeHr = hostConnections->SetConnectionSecret(
            cacheKey.c_str(), HOST_CONNECTION_SECRET_OAUTH_REFRESH_TOKEN, interactive.refreshToken.c_str(), persistRefreshToken ? TRUE : FALSE);
        if (FAILED(storeHr))
        {
            Debug::Warning(L"Microsoft Drive: Failed to store refresh token for '{}' (hr=0x{:08X}).", cacheKey.c_str(), static_cast<unsigned long>(storeHr));
        }
    }

    return cacheToken(interactive);
}

void FileSystemMicrosoftDrive::ClearCachedAccessToken(std::wstring_view connectionName) noexcept
{
    std::lock_guard lock(_stateMutex);
    auto it = _tokenCacheByConnectionName.find(std::wstring(connectionName));
    if (it == _tokenCacheByConnectionName.end())
    {
        return;
    }

    SecureClear(it->second.accessToken);
    _tokenCacheByConnectionName.erase(it);
}

void FileSystemMicrosoftDrive::UpdateDriveInfoSnapshot(std::wstring_view displayName, std::wstring_view volumeLabel) noexcept
{
    std::lock_guard lock(_stateMutex);
    _driveDisplayName.assign(displayName);
    _driveVolumeLabel.assign(volumeLabel);
}

HRESULT FileSystemMicrosoftDrive::SetPropertiesJson(std::string jsonUtf8) noexcept
{
    std::lock_guard lock(_stateMutex);
    _propertiesJson = jsonUtf8.empty() ? "{}" : std::move(jsonUtf8);
    return S_OK;
}

FileSystemMicrosoftDriveMode FileSystemMicrosoftDrive::Mode() const noexcept
{
    return _mode;
}

bool FileSystemMicrosoftDrive::TryGetCachedDrive(std::wstring_view connectionName,
                                                 std::wstring& siteIdOut,
                                                 std::wstring& driveIdOut,
                                                 std::wstring& displayNameOut,
                                                 std::wstring& volumeLabelOut,
                                                 std::wstring& webUrlOut) const noexcept
{
    std::lock_guard lock(_stateMutex);
    const auto it = _driveCacheByConnectionName.find(std::wstring(connectionName));
    if (it == _driveCacheByConnectionName.end() || it->second.driveId.empty())
    {
        return false;
    }

    siteIdOut      = it->second.siteId;
    driveIdOut     = it->second.driveId;
    displayNameOut = it->second.displayName;
    volumeLabelOut = it->second.volumeLabel;
    webUrlOut      = it->second.webUrl;
    return true;
}

void FileSystemMicrosoftDrive::StoreCachedDrive(std::wstring_view connectionName,
                                                std::wstring_view siteId,
                                                std::wstring_view driveId,
                                                std::wstring_view displayName,
                                                std::wstring_view volumeLabel,
                                                std::wstring_view webUrl) noexcept
{
    std::lock_guard lock(_stateMutex);
    CachedDrive& cached = _driveCacheByConnectionName[std::wstring(connectionName)];
    cached.siteId.assign(siteId);
    cached.driveId.assign(driveId);
    cached.displayName.assign(displayName);
    cached.volumeLabel.assign(volumeLabel);
    cached.webUrl.assign(webUrl);
}
