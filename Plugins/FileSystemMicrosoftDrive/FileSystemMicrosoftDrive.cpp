#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <winsock2.h>
#include <ws2tcpip.h>

#include <bcrypt.h>

#include <shellapi.h>
#include <winhttp.h>

#include "FileSystemMicrosoftDrive.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <charconv>
#include <chrono>
#include <cstring>
#include <format>
#include <functional>
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

#include "FileSystemMicrosoftDriveResources.h"
#include "HandleIo.h"
#include "Helpers.h"
#include "PaginationGuard.h"
#include "UriEncoding.h"
#include "YyjsonHelpers.h"
#include "resource.h"

extern HINSTANCE g_hInstance;

namespace FsMs
{
constexpr uint64_t kDefaultThrottleDelayMs    = 1'000ull;
constexpr uint64_t kAuthTimeoutMs             = 5ull * 60ull * 1'000ull;
constexpr uint64_t kGraphSimpleUploadMaxBytes = 250ull * 1024ull * 1024ull;
constexpr uint64_t kGraphChunkAlignmentBytes  = 320ull * 1024ull;

constexpr std::wstring_view kGraphBaseUrl                   = L"https://graph.microsoft.com/v1.0";
constexpr std::wstring_view kAuthHost                       = L"login.microsoftonline.com";
constexpr std::wstring_view kLoopbackPath                   = L"/redsalamander/oauth2";
constexpr std::wstring_view kDefaultAuthUserAgent           = L"RedSalamander Microsoft Drive/0.1";
constexpr std::wstring_view kDefaultOAuthPageSummarySuccess = L"Your Microsoft account is linked. Return to RedSalamander to keep browsing.";
constexpr std::wstring_view kDefaultOAuthPageSummaryFailure = L"Microsoft sign-in did not finish in this browser tab. Return to RedSalamander and try again.";
constexpr std::wstring_view kDefaultOAuthPageFunSuccess     = L"sparkles|Cloud handshake complete. Your files are ready to roam.";
constexpr std::wstring_view kDefaultOAuthPageFunFailure     = L"warning|The sign-in lost its footing. Head back for another try.";

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
    std::wstring eTag;
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
    return Common::Strings::Utf16FromUtf8StrictOrEmpty(text);
}

[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    return Common::Strings::Utf8FromUtf16StrictOrEmpty(text);
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

[[nodiscard]] std::wstring LoadEmbeddedStringResourceOrFallback(unsigned int id, std::wstring_view fallback) noexcept
{
    std::wstring text = LoadEmbeddedStringResource(g_hInstance, id);
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
            static const std::wstring text = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_SHAREPOINT_NAME);
            return text.c_str();
        }
    }

    static const std::wstring fallback = LoadEmbeddedStringResource(g_hInstance, IDS_FILESYSTEMMICROSOFTDRIVE_NAME);
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

struct AuthPageMessageVariant
{
    std::string emojiHtml;
    std::string message;
};

[[nodiscard]] std::wstring_view TrimAuthPageTextPart(std::wstring_view text) noexcept
{
    while (! text.empty())
    {
        const wchar_t ch = text.front();
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }

        text.remove_prefix(1);
    }

    while (! text.empty())
    {
        const wchar_t ch = text.back();
        if (ch != L' ' && ch != L'\t' && ch != L'\r' && ch != L'\n')
        {
            break;
        }

        text.remove_suffix(1);
    }

    return text;
}

[[nodiscard]] std::string LookupAuthPageEmojiHtml(const std::wstring_view token, const bool success) noexcept
{
    if (token == L"sparkles")
    {
        return "&#x2728;";
    }
    if (token == L"party")
    {
        return "&#x1F389;";
    }
    if (token == L"rocket")
    {
        return "&#x1F680;";
    }
    if (token == L"compass")
    {
        return "&#x1F9ED;";
    }
    if (token == L"cloud")
    {
        return "&#x2601;&#xFE0F;";
    }
    if (token == L"glowstar")
    {
        return "&#x1F31F;";
    }
    if (token == L"check")
    {
        return "&#x2705;";
    }
    if (token == L"fire")
    {
        return "&#x1F525;";
    }
    if (token == L"rainbow")
    {
        return "&#x1F308;";
    }
    if (token == L"wave")
    {
        return "&#x1F44B;";
    }
    if (token == L"warning")
    {
        return "&#x26A0;&#xFE0F;";
    }
    if (token == L"rain")
    {
        return "&#x1F327;&#xFE0F;";
    }
    if (token == L"repair")
    {
        return "&#x1F6E0;&#xFE0F;";
    }
    if (token == L"detour")
    {
        return "&#x1F9ED;";
    }
    if (token == L"anchor")
    {
        return "&#x2693;";
    }
    if (token == L"flash")
    {
        return "&#x26A1;";
    }
    if (token == L"umbrella")
    {
        return "&#x2602;&#xFE0F;";
    }
    if (token == L"map")
    {
        return "&#x1F5FA;&#xFE0F;";
    }
    if (token == L"tools")
    {
        return "&#x1F527;";
    }
    if (token == L"retry")
    {
        return "&#x1F501;";
    }

    return success ? "&#x2728;" : "&#x26A0;&#xFE0F;";
}

[[nodiscard]] AuthPageMessageVariant LoadAuthPageMessageVariant(bool success) noexcept
{
    constexpr std::array<unsigned int, 10> kSuccessMessageIds = {{
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_SUCCESS,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_SUCCESS_2,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_SUCCESS_3,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_SUCCESS_4,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_SUCCESS_5,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_SUCCESS_6,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_SUCCESS_7,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_SUCCESS_8,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_SUCCESS_9,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_SUCCESS_10,
    }};
    constexpr std::array<unsigned int, 10> kFailureMessageIds = {{
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_FAILURE,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_FAILURE_2,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_FAILURE_3,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_FAILURE_4,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_FAILURE_5,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_FAILURE_6,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_FAILURE_7,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_FAILURE_8,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_FAILURE_9,
        IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BODY_FAILURE_10,
    }};

    const auto& messageIds = success ? kSuccessMessageIds : kFailureMessageIds;
    const ULONGLONG tick   = GetTickCount64();
    const uint32_t seed    = static_cast<uint32_t>(tick ^ (tick >> 32) ^ GetCurrentProcessId() ^ GetCurrentThreadId());
    const unsigned int id  = messageIds[static_cast<size_t>(seed % static_cast<uint32_t>(messageIds.size()))];

    const std::wstring value = LoadStringResourceOrFallback(id, success ? kDefaultOAuthPageFunSuccess : kDefaultOAuthPageFunFailure);
    const std::wstring_view view(value);
    const size_t delimiter = view.find(L'|');
    if (delimiter == std::wstring_view::npos)
    {
        return {LookupAuthPageEmojiHtml({}, success), HtmlEscapeUtf8(TrimAuthPageTextPart(view))};
    }

    const std::wstring_view emojiToken = TrimAuthPageTextPart(view.substr(0, delimiter));
    const std::wstring_view message    = TrimAuthPageTextPart(view.substr(delimiter + 1));
    return {LookupAuthPageEmojiHtml(emojiToken, success), HtmlEscapeUtf8(message)};
}

[[nodiscard]] std::string LoadAuthPageSummaryHtml(bool success) noexcept
{
    const std::wstring value = LoadStringResourceOrFallback(success ? IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_SUMMARY_SUCCESS
                                                                    : IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_SUMMARY_FAILURE,
                                                            success ? kDefaultOAuthPageSummarySuccess : kDefaultOAuthPageSummaryFailure);

    const std::wstring_view summary = TrimAuthPageTextPart(value);
    const size_t delimiter          = summary.find(L". ");
    if (delimiter == std::wstring_view::npos)
    {
        return HtmlEscapeUtf8(summary);
    }

    const std::wstring_view firstLine  = TrimAuthPageTextPart(summary.substr(0, delimiter + 1));
    const std::wstring_view secondLine = TrimAuthPageTextPart(summary.substr(delimiter + 2));
    if (secondLine.empty())
    {
        return HtmlEscapeUtf8(firstLine);
    }

    std::string html = HtmlEscapeUtf8(firstLine);
    html.append("<br>");
    html.append(HtmlEscapeUtf8(secondLine));
    return html;
}

[[nodiscard]] std::string BuildAuthPageIllustration(bool success) noexcept
{
    if (success)
    {
        return R"svg(<svg viewBox="0 0 220 220" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
<defs>
<radialGradient id="rsSuccessCoinFace" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="translate(76 54) rotate(44.5) scale(156.77)">
<stop stop-color="#FFE8B5"/>
<stop offset="0.56" stop-color="#FFB04D"/>
<stop offset="1" stop-color="#ED6A2A"/>
</radialGradient>
<linearGradient id="rsSuccessCoinRim" x1="34" y1="28" x2="180" y2="192" gradientUnits="userSpaceOnUse">
<stop stop-color="#FFD38B"/>
<stop offset="0.54" stop-color="#F28734"/>
<stop offset="1" stop-color="#C84B22"/>
</linearGradient>
</defs>
<circle cx="112" cy="112" r="86" fill="#7A2F19" fill-opacity="0.16"/>
<circle cx="110" cy="108" r="84" fill="url(#rsSuccessCoinRim)"/>
<circle cx="110" cy="106" r="75" fill="url(#rsSuccessCoinFace)"/>
<circle cx="110" cy="106" r="75" stroke="#FFD99B" stroke-opacity="0.55" stroke-width="3"/>
<ellipse cx="82" cy="70" rx="44" ry="29" fill="white" fill-opacity="0.22"/>
<path d="M73 148C89 160 110 167 136 164" stroke="#B74220" stroke-opacity="0.16" stroke-width="13" stroke-linecap="round"/>
</svg>)svg";
    }

    return R"svg(<svg viewBox="0 0 220 220" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
<defs>
<radialGradient id="rsFailureCoinFace" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="translate(76 54) rotate(44.5) scale(156.77)">
<stop stop-color="#FFE0B0"/>
<stop offset="0.52" stop-color="#FF9850"/>
<stop offset="1" stop-color="#D95332"/>
</radialGradient>
<linearGradient id="rsFailureCoinRim" x1="34" y1="28" x2="180" y2="192" gradientUnits="userSpaceOnUse">
<stop stop-color="#FFC98A"/>
<stop offset="0.52" stop-color="#E26B3B"/>
<stop offset="1" stop-color="#9E331F"/>
</linearGradient>
</defs>
<circle cx="112" cy="112" r="86" fill="#661E17" fill-opacity="0.18"/>
<circle cx="110" cy="108" r="84" fill="url(#rsFailureCoinRim)"/>
<circle cx="110" cy="106" r="75" fill="url(#rsFailureCoinFace)"/>
<circle cx="110" cy="106" r="75" stroke="#FFCB96" stroke-opacity="0.42" stroke-width="3"/>
<ellipse cx="82" cy="70" rx="44" ry="29" fill="white" fill-opacity="0.18"/>
<path d="M72 149C88 160 110 165 137 161" stroke="#992C1B" stroke-opacity="0.20" stroke-width="13" stroke-linecap="round"/>
</svg>)svg";
}
[[nodiscard]] std::string BuildAuthResultHttpResponse(bool success) noexcept
{
    const std::string appTitle = HtmlEscapeUtf8(LoadEmbeddedStringResourceOrFallback(IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_APP_TITLE, L"RedSalamander"));
    const std::string brandKicker =
        HtmlEscapeUtf8(LoadStringResourceOrFallback(IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_BRAND_KICKER, L"Microsoft Drive connection"));
    const std::string title = HtmlEscapeUtf8(
        LoadStringResourceOrFallback(success ? IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_TITLE_SUCCESS : IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_TITLE_FAILURE,
                                     success ? L"Connection complete" : L"Connection interrupted"));
    const std::string summaryHtml           = LoadAuthPageSummaryHtml(success);
    const AuthPageMessageVariant funMessage = LoadAuthPageMessageVariant(success);
    const std::string illustration          = BuildAuthPageIllustration(success);
    const std::string footer                = HtmlEscapeUtf8(
        LoadStringResourceOrFallback(IDS_FILESYSTEMMICROSOFTDRIVE_OAUTH_PAGE_FOOTER_HINT, L"You can close this tab and go back to RedSalamander."));

    std::string html;
    html.reserve(8192);
    html.append("<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">");
    html.append("<title>");
    html.append(appTitle);
    html.append("</title><style>"
                ":root{color-scheme:light;}"
                "*{box-sizing:border-box;}"
                "html,body{min-height:100%;margin:0;}"
                "body{display:grid;place-items:center;padding:clamp(18px,4vw,34px);overflow:auto;color:#201913;"
                "font-family:\"Segoe UI Variable Display\",\"Segoe UI Variable Text\",\"Aptos\",\"Segoe UI\",system-ui,sans-serif;");
    html.append(success ? "background:radial-gradient(circle at 11% 16%,rgba(255,205,132,.60) 0,rgba(255,205,132,.22) 28%,transparent 52%),"
                          "radial-gradient(circle at 88% 80%,rgba(231,90,57,.16) 0,rgba(231,90,57,0) 32%,transparent 48%),"
                          "linear-gradient(155deg,#fffaf4 0%,#f6eee2 52%,#f7f5ef 100%);"
                        : "background:radial-gradient(circle at 11% 16%,rgba(255,180,146,.60) 0,rgba(255,180,146,.24) 28%,transparent 52%),"
                          "radial-gradient(circle at 88% 80%,rgba(188,47,44,.18) 0,rgba(188,47,44,0) 30%,transparent 48%),"
                          "linear-gradient(155deg,#fff8f4 0%,#f7ece4 52%,#f7f3ef 100%);");
    html.append("}"
                "body::before,body::after{content:\"\";position:fixed;pointer-events:none;border-radius:999px;filter:blur(12px);opacity:.72;}"
                "body::before{width:34vmax;height:34vmax;top:-11vmax;right:-9vmax;background:rgba(255,214,164,.24);}"
                "body::after{width:24vmax;height:24vmax;left:-8vmax;bottom:-8vmax;background:rgba(228,92,57,.12);}"
                ".card{position:relative;width:min(100%,940px);overflow:hidden;border-radius:36px;padding:28px 32px 24px;"
                "background:linear-gradient(180deg,rgba(255,255,255,.92) 0%,rgba(255,249,241,.84) 100%);"
                "border:1px solid rgba(108,73,43,.12);box-shadow:0 30px 90px rgba(91,60,31,.18),inset 0 1px 0 rgba(255,255,255,.68);"
                "backdrop-filter:blur(14px);}"
                ".hero{display:grid;grid-template-columns:220px minmax(0,1fr);column-gap:4px;align-items:start;}"
                ".art{position:relative;min-height:220px;display:grid;place-items:start end;overflow:visible;}"
                ".art::before{content:\"\";position:absolute;inset:36px 0 auto auto;width:190px;height:190px;border-radius:999px;"
                "background:radial-gradient(circle at 34% 28%,rgba(255,226,165,.34) 0,rgba(255,226,165,.16) 30%,rgba(235,110,44,.10) 58%,transparent "
                "76%);filter:blur(16px);}"
                ".art svg{position:relative;display:block;width:min(100%,212px);height:auto;filter:drop-shadow(0 18px 34px rgba(125,68,28,.18));}"
                ".copy{min-width:0;padding-top:6px;display:flex;flex-direction:column;margin-left:-34px;}"
                ".masthead{margin-bottom:18px;}"
                ".masthead-kicker{font-size:11px;font-weight:800;letter-spacing:.18em;text-transform:uppercase;color:#86614a;}"
                ".masthead-title{margin-top:2px;font-size:clamp(36px,5.2vw,60px);line-height:.90;font-weight:900;letter-spacing:-.06em;color:#23180f;}"
                ".flow{display:flex;flex-direction:column;width:min(100%,42rem);margin-left:auto;}");
    html.append(success ? ".note{background:linear-gradient(135deg,rgba(255,248,238,.94) 0%,rgba(255,243,225,.92) 54%,rgba(255,229,185,.90) 100%);border:1px "
                          "solid rgba(232,143,72,.22);}"
                          ".note::before{background:radial-gradient(circle at 100% 0%,rgba(255,208,117,.42) 0,rgba(255,208,117,0) 42%),radial-gradient(circle "
                          "at 0% 100%,rgba(231,90,57,.14) 0,rgba(231,90,57,0) 38%);}"
                          ".footer-mark{color:#bf5623;background:rgba(255,247,238,.86);}"
                        : ".note{background:linear-gradient(135deg,rgba(255,244,239,.94) 0%,rgba(255,234,225,.92) 54%,rgba(255,221,193,.88) 100%);border:1px "
                          "solid rgba(220,95,78,.22);}"
                          ".note::before{background:radial-gradient(circle at 100% 0%,rgba(255,196,113,.38) 0,rgba(255,196,113,0) 42%),radial-gradient(circle "
                          "at 0% 100%,rgba(220,95,78,.16) 0,rgba(220,95,78,0) 38%);}"
                          ".footer-mark{color:#bf4d35;background:rgba(255,244,239,.88);}");
    html.append(".title{margin:6px 0 "
                "10px;font-size:clamp(30px,4vw,46px);line-height:1.02;font-weight:850;letter-spacing:-.05em;color:#23180f;max-width:none;white-space:nowrap;}"
                ".summary{margin:0;width:100%;max-width:34rem;font-size:18px;line-height:1.75;color:#5d4c3f;}"
                ".summary br{display:block;content:\"\";margin-top:.15em;}"
                ".note{position:relative;display:grid;grid-template-columns:auto "
                "minmax(0,1fr);gap:20px;align-items:center;width:100%;margin-top:22px;padding:20px 26px;border-radius:30px;"
                "overflow:hidden;box-shadow:inset 0 1px 0 rgba(255,255,255,.55),0 18px 36px rgba(88,62,38,.08);}"
                ".note::before{content:\"\";position:absolute;inset:0;pointer-events:none;}"
                ".note>*{position:relative;z-index:1;}"
                ".note-emoji{display:grid;place-items:center;width:72px;height:72px;border-radius:22px;font-size:40px;line-height:1;"
                "font-family:\"Segoe UI Emoji\",\"Apple Color Emoji\",\"Noto Color Emoji\",\"Segoe UI Symbol\",sans-serif;"
                "background:linear-gradient(180deg,rgba(255,255,255,.96) 0%,rgba(255,245,232,.86) 100%);border:1px solid rgba(126,85,45,.10);box-shadow:0 12px "
                "24px rgba(110,71,34,.10);}"
                ".note-text{margin:0;font-size:18px;line-height:1.5;font-weight:800;color:#35261b;}"
                ".footer{display:flex;align-items:center;gap:10px;margin-top:24px;padding-top:18px;border-top:1px solid rgba(82,64,43,.10);"
                "font-size:13px;font-weight:600;color:#7c6b5e;}"
                ".footer-mark{width:28px;height:28px;border-radius:999px;display:grid;place-items:center;background:rgba(255,255,255,.72);"
                "border:1px solid rgba(82,64,43,.08);font-size:14px;}"
                "@media (max-width:700px){"
                "body{place-items:start center;}"
                ".card{padding:22px 18px 18px;border-radius:30px;}"
                ".hero{grid-template-columns:1fr;gap:12px;}"
                ".art{place-items:start center;min-height:180px;}"
                ".art::before{inset:22px auto auto 50%;width:170px;height:170px;transform:translateX(-50%);}"
                ".copy{padding-top:0;margin-left:0;}"
                ".masthead{margin-bottom:16px;}"
                ".flow{width:100%;margin-left:0;}"
                ".art svg{width:min(78%,204px);}"
                ".title{margin-top:12px;font-size:32px;}"
                ".summary{font-size:16px;}"
                ".note{gap:14px;padding:16px 18px;}"
                ".note-emoji{width:60px;height:60px;font-size:34px;}"
                ".note-text{font-size:16px;}"
                "}"
                "</style></head><body><main class=\"card\"><section class=\"hero\"><div class=\"art\" aria-hidden=\"true\">");
    html.append(illustration);
    html.append("</div><div class=\"copy\"><header class=\"masthead\"><div class=\"masthead-kicker\">");
    html.append(brandKicker);
    html.append("</div><div class=\"masthead-title\">");
    html.append(appTitle);
    html.append("</div></header><div class=\"flow\"><h1 class=\"title\">");
    html.append(title);
    html.append("</h1><p class=\"summary\">");
    html.append(summaryHtml);
    html.append("</p></div></div></section><section class=\"note\"><div class=\"note-emoji\" aria-hidden=\"true\">");
    html.append(funMessage.emojiHtml);
    html.append("</div><div><p class=\"note-text\">");
    html.append(funMessage.message);
    html.append("</p></div></section><div class=\"footer\"><div class=\"footer-mark\" aria-hidden=\"true\">");
    html.append(success ? "↩" : "↺");
    html.append("</div><div>");
    html.append(footer);
    html.append("</div></div></main></body></html>");

    std::string response;
    response.reserve(html.size() + 320);
    response.append(success ? "HTTP/1.1 200 OK\r\n" : "HTTP/1.1 400 Bad Request\r\n");
    response.append("Content-Type: text/html; charset=utf-8\r\n");
    response.append("Cache-Control: no-store\r\n");
    response.append("Content-Security-Policy: default-src 'none'; style-src 'unsafe-inline'; img-src data:; base-uri 'none'; form-action 'none'\r\n");
    response.append("Connection: close\r\n\r\n");
    response.append(html);
    return response;
}

using SecureWipe::SecureClear;

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

[[nodiscard]] bool IsSameOrDescendantPath(std::wstring_view ancestorPath, std::wstring_view candidatePath) noexcept
{
    if (OrdinalString::EqualsNoCase(ancestorPath, candidatePath))
    {
        return true;
    }

    if (ancestorPath == L"/")
    {
        return ! candidatePath.empty() && candidatePath != L"/";
    }

    if (candidatePath.size() <= ancestorPath.size() || candidatePath[ancestorPath.size()] != L'/')
    {
        return false;
    }

    return OrdinalString::EqualsNoCase(std::wstring(candidatePath.substr(0, ancestorPath.size())), std::wstring(ancestorPath));
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
    const Common::Json::MemberResult<std::string_view> value = Common::Json::GetStringMember(root, key, Common::Json::MemberRequirement::Optional);
    return value.HasValue() ? std::optional<std::wstring>{Utf16FromUtf8(value.value)} : std::nullopt;
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
    const Common::Json::MemberResult<bool> value = Common::Json::GetBoolMember(root, key, Common::Json::MemberRequirement::Optional);
    return value.HasValue() ? std::optional<bool>{value.value} : std::nullopt;
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
    return Common::HandleIo::Rewind(file);
}

[[nodiscard]] HRESULT GetFileSizeBytes(HANDLE file, uint64_t& sizeBytesOut) noexcept
{
    return Common::HandleIo::GetFileSizeBounded(file, (std::numeric_limits<uint64_t>::max)(), sizeBytesOut);
}

[[nodiscard]] std::wstring PercentEncodeUtf8(std::wstring_view text) noexcept
{
    std::wstring encoded;
    return Common::Uri::TryPercentEncodeUtf8ToWide(text, Common::Uri::SlashPolicy::Encode, encoded) ? encoded : std::wstring{};
}

[[nodiscard]] std::wstring PercentEncodeGraphPath(std::wstring_view path) noexcept
{
    const auto segments = SplitPathSegments(path);
    if (segments.empty())
    {
        return {};
    }

    std::wstring normalized;
    bool first = true;
    for (const std::wstring_view segment : segments)
    {
        if (! first)
        {
            normalized.push_back(L'/');
        }
        first = false;
        normalized.append(segment);
    }
    std::wstring encoded;
    return Common::Uri::TryPercentEncodeUtf8ToWide(normalized, Common::Uri::SlashPolicy::Preserve, encoded) ? encoded : std::wstring{};
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
    HRESULT hr = Common::Crypto::GenerateRandomBytes(verifierBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    std::array<std::byte, 16> stateBytes{};
    hr = Common::Crypto::GenerateRandomBytes(stateBytes);
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

[[nodiscard]] bool IsRetryableGraphThrottleStatus(DWORD statusCode) noexcept
{
    return statusCode == 429u || statusCode == 503u || statusCode == 504u;
}

[[nodiscard]] uint64_t ComputeGraphRetryDelayMs(std::wstring_view retryAfter, int attempt) noexcept
{
    uint64_t delayMs = kDefaultThrottleDelayMs;
    if (! ParseRetryAfterMs(retryAfter, delayMs))
    {
        delayMs = kDefaultThrottleDelayMs * static_cast<uint64_t>(attempt + 1);
    }
    return delayMs;
}

[[nodiscard]] bool ShouldRetryGraphHttpResponse(bool allowRetry, DWORD statusCode, int attempt) noexcept
{
    return allowRetry && IsRetryableGraphThrottleStatus(statusCode) && attempt < 3;
}

#if defined(_DEBUG)
std::atomic_bool g_debugMicrosoftDriveBypassAccessTokenForSelfTest{false};
std::atomic_bool g_debugMicrosoftDriveSuppressRetrySleepForSelfTest{false};
std::atomic_bool g_debugMicrosoftDriveUseSyntheticContextForSelfTest{false};

using DebugHttpRequestHook = HRESULT (*)(void* cookie,
                                         std::wstring_view method,
                                         std::wstring_view url,
                                         std::span<const HttpHeader> headers,
                                         std::string_view bodyUtf8,
                                         bool allowRetry,
                                         HttpResponse& responseOut) noexcept;

std::mutex g_debugMicrosoftDriveHttpHookMutex;
DebugHttpRequestHook g_debugMicrosoftDriveHttpHook = nullptr;
void* g_debugMicrosoftDriveHttpHookCookie          = nullptr;

using DebugHttpDiagnosticHook                                   = void (*)(void* cookie, std::wstring_view diagnostic) noexcept;
DebugHttpDiagnosticHook g_debugMicrosoftDriveHttpDiagnosticHook = nullptr;
void* g_debugMicrosoftDriveHttpDiagnosticHookCookie             = nullptr;

class DebugHttpRequestHookScope final
{
public:
    DebugHttpRequestHookScope(DebugHttpRequestHook hook, void* cookie) noexcept
    {
        std::lock_guard lock(g_debugMicrosoftDriveHttpHookMutex);
        _previousHook                       = g_debugMicrosoftDriveHttpHook;
        _previousCookie                     = g_debugMicrosoftDriveHttpHookCookie;
        g_debugMicrosoftDriveHttpHook       = hook;
        g_debugMicrosoftDriveHttpHookCookie = cookie;
    }

    ~DebugHttpRequestHookScope()
    {
        std::lock_guard lock(g_debugMicrosoftDriveHttpHookMutex);
        g_debugMicrosoftDriveHttpHook       = _previousHook;
        g_debugMicrosoftDriveHttpHookCookie = _previousCookie;
    }

    DebugHttpRequestHookScope(const DebugHttpRequestHookScope&)            = delete;
    DebugHttpRequestHookScope& operator=(const DebugHttpRequestHookScope&) = delete;
    DebugHttpRequestHookScope(DebugHttpRequestHookScope&&)                 = delete;
    DebugHttpRequestHookScope& operator=(DebugHttpRequestHookScope&&)      = delete;

private:
    DebugHttpRequestHook _previousHook = nullptr;
    void* _previousCookie              = nullptr;
};

class DebugHttpDiagnosticHookScope final
{
public:
    DebugHttpDiagnosticHookScope(DebugHttpDiagnosticHook hook, void* cookie) noexcept
    {
        std::lock_guard lock(g_debugMicrosoftDriveHttpHookMutex);
        _previousHook                                 = g_debugMicrosoftDriveHttpDiagnosticHook;
        _previousCookie                               = g_debugMicrosoftDriveHttpDiagnosticHookCookie;
        g_debugMicrosoftDriveHttpDiagnosticHook       = hook;
        g_debugMicrosoftDriveHttpDiagnosticHookCookie = cookie;
    }

    ~DebugHttpDiagnosticHookScope()
    {
        std::lock_guard lock(g_debugMicrosoftDriveHttpHookMutex);
        g_debugMicrosoftDriveHttpDiagnosticHook       = _previousHook;
        g_debugMicrosoftDriveHttpDiagnosticHookCookie = _previousCookie;
    }

    DebugHttpDiagnosticHookScope(const DebugHttpDiagnosticHookScope&)            = delete;
    DebugHttpDiagnosticHookScope& operator=(const DebugHttpDiagnosticHookScope&) = delete;
    DebugHttpDiagnosticHookScope(DebugHttpDiagnosticHookScope&&)                 = delete;
    DebugHttpDiagnosticHookScope& operator=(DebugHttpDiagnosticHookScope&&)      = delete;

private:
    DebugHttpDiagnosticHook _previousHook = nullptr;
    void* _previousCookie                 = nullptr;
};

class DebugFlagScope final
{
public:
    DebugFlagScope(std::atomic_bool& flag, bool value) noexcept : _flag(flag), _previous(flag.exchange(value))
    {
    }

    ~DebugFlagScope()
    {
        _flag.store(_previous);
    }

    DebugFlagScope(const DebugFlagScope&)            = delete;
    DebugFlagScope& operator=(const DebugFlagScope&) = delete;
    DebugFlagScope(DebugFlagScope&&)                 = delete;
    DebugFlagScope& operator=(DebugFlagScope&&)      = delete;

private:
    std::atomic_bool& _flag;
    bool _previous = false;
};

[[nodiscard]] bool TryHandleDebugHttpRequest(std::wstring_view method,
                                             std::wstring_view url,
                                             std::span<const HttpHeader> headers,
                                             const std::byte* bodyBytes,
                                             size_t bodySizeBytes,
                                             bool allowRetry,
                                             HttpResponse& responseOut,
                                             HRESULT& hrOut) noexcept
{
    DebugHttpRequestHook hook = nullptr;
    void* cookie              = nullptr;
    {
        std::lock_guard lock(g_debugMicrosoftDriveHttpHookMutex);
        hook   = g_debugMicrosoftDriveHttpHook;
        cookie = g_debugMicrosoftDriveHttpHookCookie;
    }

    if (! hook)
    {
        return false;
    }

    const std::string_view bodyUtf8 =
        bodyBytes && bodySizeBytes != 0u ? std::string_view(reinterpret_cast<const char*>(bodyBytes), bodySizeBytes) : std::string_view{};
    hrOut = hook(cookie, method, url, headers, bodyUtf8, allowRetry, responseOut);
    return true;
}
#endif

[[nodiscard]] std::wstring DescribeHttpRequestTarget(std::wstring_view url) noexcept;

void SleepBeforeGraphRetry(uint64_t delayMs) noexcept
{
#if defined(_DEBUG)
    if (g_debugMicrosoftDriveSuppressRetrySleepForSelfTest.load())
    {
        return;
    }
#endif
    std::this_thread::sleep_for(std::chrono::milliseconds(delayMs));
}

void LogAndSleepBeforeGraphRetry(std::wstring_view method, std::wstring_view url, DWORD statusCode, std::wstring_view retryAfter, int attempt) noexcept
{
    const uint64_t delayMs = ComputeGraphRetryDelayMs(retryAfter, attempt);
    Debug::Info(L"Microsoft Drive: retrying request after throttle. method='{}' target='{}' status={} delayMs={}",
                method,
                DescribeHttpRequestTarget(url),
                statusCode,
                delayMs);
    SleepBeforeGraphRetry(delayMs);
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

        SessionState()                               = default;
        SessionState(const SessionState&)            = delete;
        SessionState& operator=(const SessionState&) = delete;
        SessionState(SessionState&&)                 = delete;
        SessionState& operator=(SessionState&&)      = delete;
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

struct ParsedHttpUrl final
{
    ParsedHttpUrl() = default;
    ~ParsedHttpUrl()
    {
        SecureClear(pathAndQuery);
    }

    ParsedHttpUrl(const ParsedHttpUrl&)            = delete;
    ParsedHttpUrl& operator=(const ParsedHttpUrl&) = delete;

    INTERNET_SCHEME scheme = static_cast<INTERNET_SCHEME>(0);
    INTERNET_PORT port     = 0;
    std::wstring host;
    std::wstring pathAndQuery;
    bool hasUserInfo = false;
    bool hasFragment = false;
};

struct ValidatedGraphApiUrl final
{
    ValidatedGraphApiUrl() = default;
    ~ValidatedGraphApiUrl()
    {
        SecureClear(value);
    }

    ValidatedGraphApiUrl(const ValidatedGraphApiUrl&)            = delete;
    ValidatedGraphApiUrl& operator=(const ValidatedGraphApiUrl&) = delete;

    std::wstring value;
};

struct ValidatedPreauthenticatedUploadUrl final
{
    ValidatedPreauthenticatedUploadUrl() = default;
    ~ValidatedPreauthenticatedUploadUrl()
    {
        SecureClear(value);
    }

    ValidatedPreauthenticatedUploadUrl(const ValidatedPreauthenticatedUploadUrl&)            = delete;
    ValidatedPreauthenticatedUploadUrl& operator=(const ValidatedPreauthenticatedUploadUrl&) = delete;

    std::wstring value;

    [[nodiscard]] bool empty() const noexcept
    {
        return value.empty();
    }
};

[[nodiscard]] HRESULT CrackUrl(std::wstring_view url, ParsedHttpUrl& parsedOut) noexcept
{
    parsedOut.scheme = static_cast<INTERNET_SCHEME>(0);
    parsedOut.port   = 0;
    parsedOut.host.clear();
    SecureClear(parsedOut.pathAndQuery);
    parsedOut.hasUserInfo = false;
    parsedOut.hasFragment = false;

    std::wstring urlCopy(url);
    URL_COMPONENTS components{};
    components.dwStructSize      = sizeof(components);
    components.dwSchemeLength    = static_cast<DWORD>(-1);
    components.dwHostNameLength  = static_cast<DWORD>(-1);
    components.dwUserNameLength  = static_cast<DWORD>(-1);
    components.dwPasswordLength  = static_cast<DWORD>(-1);
    components.dwUrlPathLength   = static_cast<DWORD>(-1);
    components.dwExtraInfoLength = static_cast<DWORD>(-1);

    // We want pointers back into urlCopy; ICU_ESCAPE is only valid when WinHTTP copies into caller buffers.
    if (WinHttpCrackUrl(urlCopy.data(), static_cast<DWORD>(urlCopy.size()), 0, &components) == 0)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    parsedOut.scheme = components.nScheme;
    parsedOut.port   = components.nPort;
    parsedOut.host.assign(components.lpszHostName, components.dwHostNameLength);
    parsedOut.pathAndQuery.assign(components.lpszUrlPath, components.dwUrlPathLength);
    if (components.dwExtraInfoLength > 0)
    {
        parsedOut.pathAndQuery.append(components.lpszExtraInfo, components.dwExtraInfoLength);
    }

    if (parsedOut.pathAndQuery.empty())
    {
        parsedOut.pathAndQuery = L"/";
    }

    parsedOut.hasUserInfo = components.dwUserNameLength != 0u || components.dwPasswordLength != 0u;
    parsedOut.hasFragment = url.find(L'#') != std::wstring_view::npos;

    return S_OK;
}

[[nodiscard]] bool IsHostOrSubdomainOf(std::wstring_view host, std::wstring_view approvedRoot) noexcept
{
    if (OrdinalString::EqualsNoCase(host, approvedRoot))
    {
        return true;
    }
    if (host.size() <= approvedRoot.size() || host[host.size() - approvedRoot.size() - 1u] != L'.')
    {
        return false;
    }
    return OrdinalString::EqualsNoCase(host.substr(host.size() - approvedRoot.size()), approvedRoot);
}

[[nodiscard]] bool IsApprovedPreauthenticatedUploadHost(std::wstring_view host) noexcept
{
    // Microsoft Graph documents OneDrive upload-session URLs under *.up.1drv.com. Business and
    // SharePoint sessions use the tenant's *.sharepoint.com service origin.
    return IsHostOrSubdomainOf(host, L"up.1drv.com") || IsHostOrSubdomainOf(host, L"sharepoint.com");
}

[[nodiscard]] bool IsGraphV1Path(std::wstring_view pathAndQuery) noexcept
{
    const size_t queryOffset     = pathAndQuery.find(L'?');
    const std::wstring_view path = queryOffset == std::wstring_view::npos ? pathAndQuery : pathAndQuery.substr(0u, queryOffset);
    return path == L"/v1.0" || (path.size() > 6u && path.starts_with(L"/v1.0/"));
}

[[nodiscard]] HRESULT ValidateGraphApiUrl(std::wstring_view url, ValidatedGraphApiUrl& validatedOut) noexcept
{
    SecureClear(validatedOut.value);
    ParsedHttpUrl parsed{};
    const HRESULT hr = CrackUrl(url, parsed);
    if (FAILED(hr))
    {
        return hr;
    }
    if (parsed.scheme != INTERNET_SCHEME_HTTPS || parsed.port != INTERNET_DEFAULT_HTTPS_PORT || parsed.hasUserInfo || parsed.hasFragment ||
        ! OrdinalString::EqualsNoCase(parsed.host, L"graph.microsoft.com") || ! IsGraphV1Path(parsed.pathAndQuery))
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
    validatedOut.value.assign(url);
    return S_OK;
}

[[nodiscard]] HRESULT ValidatePreauthenticatedUploadUrl(std::wstring_view url, ValidatedPreauthenticatedUploadUrl& validatedOut) noexcept
{
    SecureClear(validatedOut.value);
    ParsedHttpUrl parsed{};
    const HRESULT hr = CrackUrl(url, parsed);
    if (FAILED(hr))
    {
        return hr;
    }
    if (parsed.scheme != INTERNET_SCHEME_HTTPS || parsed.port != INTERNET_DEFAULT_HTTPS_PORT || parsed.hasUserInfo || parsed.hasFragment ||
        ! IsApprovedPreauthenticatedUploadHost(parsed.host))
    {
        return HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED);
    }
    validatedOut.value.assign(url);
    return S_OK;
}

[[nodiscard]] std::wstring DescribeHttpRequestTarget(std::wstring_view url) noexcept
{
    ParsedHttpUrl parsed{};
    if (FAILED(CrackUrl(url, parsed)))
    {
        return L"invalid";
    }

    const bool https = parsed.scheme == INTERNET_SCHEME_HTTPS;
    if (https && OrdinalString::EqualsNoCase(parsed.host, L"graph.microsoft.com") && IsGraphV1Path(parsed.pathAndQuery))
    {
        return L"https/graph-api";
    }
    if (https && OrdinalString::EqualsNoCase(parsed.host, kAuthHost))
    {
        return L"https/oauth-authority";
    }
    if (https && IsApprovedPreauthenticatedUploadHost(parsed.host))
    {
        return L"https/preauthenticated-upload";
    }
    return https ? L"https/external" : L"plaintext-or-unknown/external";
}

void EmitSanitizedHttpDiagnostic(std::wstring diagnostic) noexcept
{
#if defined(_DEBUG)
    DebugHttpDiagnosticHook hook = nullptr;
    void* cookie                 = nullptr;
    {
        std::lock_guard lock(g_debugMicrosoftDriveHttpHookMutex);
        hook   = g_debugMicrosoftDriveHttpDiagnosticHook;
        cookie = g_debugMicrosoftDriveHttpDiagnosticHookCookie;
    }
    if (hook)
    {
        hook(cookie, diagnostic);
    }
#endif
    Debug::Warning(L"{}", diagnostic);
}

void LogHttpTransportFailure(
    std::wstring_view stage, std::wstring_view method, std::wstring_view url, HRESULT hr, size_t byteCount, size_t headerCount, bool bearerPresent) noexcept
{
    EmitSanitizedHttpDiagnostic(std::format(L"Microsoft Drive: HTTP request failed. stage='{}' method='{}' target='{}' hr=0x{:08X} bytes={} "
                                            L"headerCount={} bearerPresent={}.",
                                            stage,
                                            method,
                                            DescribeHttpRequestTarget(url),
                                            static_cast<unsigned long>(hr),
                                            byteCount,
                                            headerCount,
                                            bearerPresent));
}

void LogHttpResponseFailure(std::wstring_view method, std::wstring_view url, DWORD status, std::wstring_view requestId, size_t responseBytes) noexcept
{
    Debug::Warning(L"Microsoft Drive: HTTP response failed. method='{}' target='{}' status={} requestId='{}' bytes={}.",
                   method,
                   DescribeHttpRequestTarget(url),
                   status,
                   requestId.empty() ? std::wstring_view(L"<none>") : requestId,
                   responseBytes);
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
                                      bool disableRedirects,
                                      HttpResponse& responseOut) noexcept
{
    responseOut              = {};
    const bool bearerPresent = bearerToken && bearerToken[0] != '\0';

    const HINTERNET session = GetSharedWinHttpSession();
    if (! session)
    {
        return HRESULT_FROM_WIN32(ERROR_WINHTTP_INTERNAL_ERROR);
    }

    for (int attempt = 0; attempt < 4; ++attempt)
    {
#if defined(_DEBUG)
        {
            HRESULT debugHr = S_OK;
            if (TryHandleDebugHttpRequest(method, url, headers, bodyBytes, bodySizeBytes, allowRetry, responseOut, debugHr))
            {
                if (FAILED(debugHr))
                {
                    LogHttpTransportFailure(L"debug-transport", method, url, debugHr, bodySizeBytes, headers.size(), bearerPresent);
                    return debugHr;
                }
                if (ShouldRetryGraphHttpResponse(allowRetry, responseOut.statusCode, attempt))
                {
                    LogAndSleepBeforeGraphRetry(method, url, responseOut.statusCode, responseOut.retryAfter, attempt);
                    continue;
                }
                return S_OK;
            }
        }
#endif

        ParsedHttpUrl parsed{};
        HRESULT hr = CrackUrl(url, parsed);
        if (FAILED(hr))
        {
            LogHttpTransportFailure(L"CrackUrl", method, url, hr, bodySizeBytes, headers.size(), bearerPresent);
            return hr;
        }

        wil::unique_winhttp_hinternet connection(WinHttpConnect(session, parsed.host.c_str(), parsed.port, 0));
        if (! connection)
        {
            const DWORD lastError = GetLastError();
            LogHttpTransportFailure(L"WinHttpConnect", method, url, HRESULT_FROM_WIN32(lastError), bodySizeBytes, headers.size(), bearerPresent);
            return HRESULT_FROM_WIN32(lastError);
        }

        const DWORD requestFlags = (parsed.scheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
        wil::unique_winhttp_hinternet request(WinHttpOpenRequest(connection.get(),
                                                                 std::wstring(method).c_str(),
                                                                 parsed.pathAndQuery.c_str(),
                                                                 nullptr,
                                                                 WINHTTP_NO_REFERER,
                                                                 WINHTTP_DEFAULT_ACCEPT_TYPES,
                                                                 requestFlags));
        if (! request)
        {
            const DWORD lastError = GetLastError();
            LogHttpTransportFailure(L"WinHttpOpenRequest", method, url, HRESULT_FROM_WIN32(lastError), bodySizeBytes, headers.size(), bearerPresent);
            return HRESULT_FROM_WIN32(lastError);
        }

        if (disableRedirects)
        {
            DWORD disabledFeatures = WINHTTP_DISABLE_REDIRECTS;
            if (WinHttpSetOption(request.get(), WINHTTP_OPTION_DISABLE_FEATURE, &disabledFeatures, sizeof(disabledFeatures)) == 0)
            {
                const DWORD lastError = GetLastError();
                LogHttpTransportFailure(
                    L"WinHttpSetOption(disable-redirects)", method, url, HRESULT_FROM_WIN32(lastError), bodySizeBytes, headers.size(), bearerPresent);
                return HRESULT_FROM_WIN32(lastError);
            }
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
        auto clearAuthorizationHeader = wil::scope_exit([&] { SecureClear(authorizationHeader); });
        if (bearerPresent)
        {
            authorizationHeader = L"Authorization: Bearer ";
            authorizationHeader.append(Utf16FromUtf8(bearerToken));
            if (WinHttpAddRequestHeaders(request.get(),
                                         authorizationHeader.c_str(),
                                         static_cast<DWORD>(authorizationHeader.size()),
                                         WINHTTP_ADDREQ_FLAG_ADD | WINHTTP_ADDREQ_FLAG_REPLACE) == 0)
            {
                const DWORD lastError = GetLastError();
                LogHttpTransportFailure(L"WinHttpAddRequestHeaders", method, url, HRESULT_FROM_WIN32(lastError), bodySizeBytes, headers.size(), bearerPresent);
                return HRESULT_FROM_WIN32(lastError);
            }
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
            LogHttpTransportFailure(L"WinHttpSendRequest", method, url, HRESULT_FROM_WIN32(lastError), bodySizeBytes, headers.size(), bearerPresent);
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
                    LogHttpTransportFailure(L"ReadFile(request-body)", method, url, HRESULT_FROM_WIN32(lastError), read, headers.size(), bearerPresent);
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
                    LogHttpTransportFailure(L"WinHttpWriteData", method, url, HRESULT_FROM_WIN32(lastError), read, headers.size(), bearerPresent);
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
            LogHttpTransportFailure(L"WinHttpReceiveResponse", method, url, HRESULT_FROM_WIN32(lastError), bodySizeBytes, headers.size(), bearerPresent);
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
            LogHttpTransportFailure(L"WinHttpQueryHeaders(status)", method, url, HRESULT_FROM_WIN32(lastError), 0u, headers.size(), bearerPresent);
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
                LogHttpTransportFailure(
                    L"WinHttpQueryDataAvailable", method, url, HRESULT_FROM_WIN32(lastError), responseOut.body.size(), headers.size(), bearerPresent);
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
                LogHttpTransportFailure(L"WinHttpReadData", method, url, HRESULT_FROM_WIN32(lastError), available, headers.size(), bearerPresent);
                return HRESULT_FROM_WIN32(lastError);
            }
            responseOut.body.resize(start + read);
        }

        if (ShouldRetryGraphHttpResponse(allowRetry, statusCode, attempt))
        {
            LogAndSleepBeforeGraphRetry(method, url, statusCode, responseOut.retryAfter, attempt);
            continue;
        }

        return S_OK;
    }

    return E_FAIL;
}

[[nodiscard]] HRESULT SendAuthenticatedGraphHttpRequest(const FileSystemMicrosoftDrive::Settings& settings,
                                                        std::wstring_view method,
                                                        std::wstring_view rawUrl,
                                                        const char* bearerToken,
                                                        std::span<const HttpHeader> headers,
                                                        const std::byte* bodyBytes,
                                                        size_t bodySizeBytes,
                                                        HANDLE bodyFile,
                                                        bool allowRetry,
                                                        HttpResponse& responseOut) noexcept
{
    ValidatedGraphApiUrl graphUrl{};
    const HRESULT hr = ValidateGraphApiUrl(rawUrl, graphUrl);
    if (FAILED(hr))
    {
        LogHttpTransportFailure(L"ValidateGraphApiUrl", method, rawUrl, hr, bodySizeBytes, headers.size(), bearerToken && bearerToken[0] != '\0');
        return hr;
    }
    return SendHttpRequest(settings, method, graphUrl.value, bearerToken, headers, bodyBytes, bodySizeBytes, bodyFile, allowRetry, true, responseOut);
}

[[nodiscard]] HRESULT SendPreauthenticatedUploadRequest(const FileSystemMicrosoftDrive::Settings& settings,
                                                        std::wstring_view method,
                                                        const ValidatedPreauthenticatedUploadUrl& uploadUrl,
                                                        std::span<const HttpHeader> headers,
                                                        const std::byte* bodyBytes,
                                                        size_t bodySizeBytes,
                                                        bool allowRetry,
                                                        HttpResponse& responseOut) noexcept
{
    if (uploadUrl.empty())
    {
        return E_INVALIDARG;
    }
    return SendHttpRequest(settings, method, uploadUrl.value, nullptr, headers, bodyBytes, bodySizeBytes, nullptr, allowRetry, true, responseOut);
}

[[nodiscard]] HRESULT HresultFromGraphError(DWORD statusCode, std::string_view bodyUtf8) noexcept
{
    std::wstring code;
    if (! bodyUtf8.empty())
    {
        yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
        if (doc)
        {
            Common::Json::UniqueDocument docOwner{doc};
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
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
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
    Common::Json::UniqueDocument docOwner{doc};

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
        settings, L"POST", url, nullptr, headers, reinterpret_cast<const std::byte*>(formUtf8.data()), formUtf8.size(), nullptr, true, false, response);
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
    Common::Json::UniqueDocument docOwner{doc};

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
    Common::Json::UniqueDocument docOwner{doc};

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
    itemOut.eTag           = TryGetJsonString(root, "eTag").value_or(L"");
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
    Common::Json::UniqueDocument docOwner{doc};

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
    Common::Json::UniqueDocument docOwner{doc};

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
                                    std::wstring& nextLinkOut,
                                    bool* incompleteDueToInvalidChildNameOut = nullptr) noexcept
{
    entriesOut.clear();
    nextLinkOut.clear();
    if (incompleteDueToInvalidChildNameOut)
    {
        *incompleteDueToInvalidChildNameOut = false;
    }

    yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
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
            // A malformed (non-object) array element means we cannot account for this child. Mark the
            // enumeration incomplete so a recursive merge-move never treats the source as fully drained
            // and deletes it -- same data-safety contract as the empty-name case below.
            if (incompleteDueToInvalidChildNameOut)
            {
                *incompleteDueToInvalidChildNameOut = true;
            }
            continue;
        }

        FilesInformationMicrosoftDrive::Entry entry{};
        entry.name = TryGetJsonString(item, "name").value_or(L"");
        if (entry.name.empty())
        {
            if (incompleteDueToInvalidChildNameOut)
            {
                *incompleteDueToInvalidChildNameOut = true;
            }
            continue;
        }
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

[[nodiscard]] std::string FormatItemPropertiesSize(uint64_t sizeBytes)
{
    const std::wstring exactBytes = std::format(L"{} bytes", sizeBytes);
    if (sizeBytes < 1024ull)
    {
        return Utf8FromUtf16(exactBytes);
    }

    return Utf8FromUtf16(std::format(L"{} ({})", FormatBytesCompact(sizeBytes), exactBytes));
}

[[nodiscard]] std::string FormatFileTimeLocal(__int64 fileTimeValue) noexcept
{
    if (fileTimeValue <= 0)
    {
        return {};
    }

    ULARGE_INTEGER ull{};
    ull.QuadPart = static_cast<ULONGLONG>(fileTimeValue);

    FILETIME fileTime{};
    fileTime.dwLowDateTime  = ull.LowPart;
    fileTime.dwHighDateTime = ull.HighPart;

    FILETIME localFileTime{};
    if (FileTimeToLocalFileTime(&fileTime, &localFileTime) == 0)
    {
        return {};
    }

    SYSTEMTIME localSystemTime{};
    if (FileTimeToSystemTime(&localFileTime, &localSystemTime) == 0)
    {
        return {};
    }

    return Utf8FromUtf16(std::format(L"{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                                     localSystemTime.wYear,
                                     localSystemTime.wMonth,
                                     localSystemTime.wDay,
                                     localSystemTime.wHour,
                                     localSystemTime.wMinute,
                                     localSystemTime.wSecond));
}

[[nodiscard]] std::string JsonScalarToStringUtf8(yyjson_val* value)
{
    if (! value)
    {
        return {};
    }

    if (yyjson_is_str(value))
    {
        const char* text = yyjson_get_str(value);
        return text ? std::string(text, yyjson_get_len(value)) : std::string();
    }
    if (yyjson_is_bool(value))
    {
        return yyjson_get_bool(value) != 0 ? "true" : "false";
    }
    if (yyjson_is_null(value))
    {
        return "null";
    }
    if (yyjson_is_uint(value))
    {
        return std::format("{}", yyjson_get_uint(value));
    }
    if (yyjson_is_sint(value))
    {
        return std::format("{}", yyjson_get_sint(value));
    }
    if (yyjson_is_real(value))
    {
        return std::format("{}", yyjson_get_real(value));
    }

    return {};
}

void AddPropertiesField(yyjson_mut_doc* doc, yyjson_mut_val* fields, std::string_view key, std::string_view value) noexcept
{
    if (! doc || ! fields || key.empty() || value.empty())
    {
        return;
    }

    yyjson_mut_val* field = yyjson_mut_obj(doc);
    yyjson_mut_obj_add_strncpy(doc, field, "key", key.data(), key.size());
    yyjson_mut_obj_add_strncpy(doc, field, "value", value.data(), value.size());
    yyjson_mut_arr_add_val(fields, field);
}

void AddPropertiesFieldWide(yyjson_mut_doc* doc, yyjson_mut_val* fields, std::string_view key, std::wstring_view value) noexcept
{
    if (value.empty())
    {
        return;
    }

    AddPropertiesField(doc, fields, key, Utf8FromUtf16(value));
}

[[nodiscard]] yyjson_mut_val* AddPropertiesSection(yyjson_mut_doc* doc, yyjson_mut_val* sections, const char* title) noexcept
{
    yyjson_mut_val* section = yyjson_mut_obj(doc);
    yyjson_mut_obj_add_strcpy(doc, section, "title", title);

    yyjson_mut_val* fields = yyjson_mut_arr(doc);
    yyjson_mut_obj_add_val(doc, section, "fields", fields);
    yyjson_mut_arr_add_val(sections, section);
    return fields;
}

void AddJsonScalarFields(yyjson_mut_doc* doc, yyjson_mut_val* fields, yyjson_val* value, std::string_view prefix, size_t depth) noexcept
{
    if (! doc || ! fields || ! value || depth > 3u)
    {
        return;
    }

    if (yyjson_is_obj(value))
    {
        size_t index       = 0u;
        size_t max         = 0u;
        yyjson_val* keyVal = nullptr;
        yyjson_val* val    = nullptr;
        yyjson_obj_foreach(value, index, max, keyVal, val)
        {
            const char* keyText = yyjson_get_str(keyVal);
            if (! keyText)
            {
                continue;
            }

            const std::string key(keyText, yyjson_get_len(keyVal));
            if (key == "@microsoft.graph.downloadUrl")
            {
                continue;
            }

            std::string displayKey(prefix);
            if (! displayKey.empty())
            {
                displayKey.push_back('.');
            }
            displayKey.append(key);

            AddJsonScalarFields(doc, fields, val, displayKey, depth + 1u);
        }
        return;
    }

    if (yyjson_is_arr(value))
    {
        AddPropertiesField(doc, fields, std::string(prefix) + ".count", std::format("{}", yyjson_arr_size(value)));
        return;
    }

    AddPropertiesField(doc, fields, prefix, JsonScalarToStringUtf8(value));
}

[[nodiscard]] HRESULT BuildItemPropertiesJson(const DriveContext& context, const ItemMetadata& item, std::string& jsonUtf8Out) noexcept
{
    yyjson_doc* rawDoc = nullptr;
    if (! item.rawJson.empty())
    {
        rawDoc = yyjson_read(item.rawJson.data(), item.rawJson.size(), YYJSON_READ_ALLOW_BOM);
    }
    Common::Json::UniqueDocument rawDocOwner{rawDoc};

    yyjson_val* rawRoot = rawDoc ? yyjson_doc_get_root(rawDoc) : nullptr;
    if (rawRoot && ! yyjson_is_obj(rawRoot))
    {
        rawRoot = nullptr;
    }

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }
    Common::Json::UniqueMutableDocument docOwner{doc};

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    yyjson_mut_doc_set_root(doc, root);
    yyjson_mut_obj_add_int(doc, root, "version", 1);
    yyjson_mut_obj_add_strcpy(doc, root, "title", "properties");

    yyjson_mut_val* sections = yyjson_mut_arr(doc);
    yyjson_mut_obj_add_val(doc, root, "sections", sections);

    const std::wstring normalizedPath = NormalizePluginPath(context.drivePath);
    const bool isDirectory            = item.isFolder || (item.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;

    yyjson_mut_val* general = AddPropertiesSection(doc, sections, "General");
    AddPropertiesFieldWide(doc, general, "Name", item.name.empty() ? std::wstring_view(normalizedPath) : std::wstring_view(item.name));
    AddPropertiesFieldWide(doc, general, "Path", normalizedPath);
    AddPropertiesField(doc, general, "Type", isDirectory ? "Directory" : "File");
    if (! isDirectory || item.sizeBytes != 0u)
    {
        AddPropertiesField(doc, general, "Size", FormatItemPropertiesSize(item.sizeBytes));
    }

    yyjson_mut_val* remote = AddPropertiesSection(doc, sections, "Remote");
    AddPropertiesFieldWide(doc, remote, "Item ID", item.id);
    AddPropertiesFieldWide(doc, remote, "Drive path", normalizedPath);
    if (rawRoot)
    {
        AddPropertiesFieldWide(doc, remote, "Web URL", TryGetJsonString(rawRoot, "webUrl").value_or(L""));
        AddPropertiesFieldWide(doc, remote, "ETag", TryGetJsonString(rawRoot, "eTag").value_or(L""));
        AddPropertiesFieldWide(doc, remote, "CTag", TryGetJsonString(rawRoot, "cTag").value_or(L""));
    }
    AddPropertiesField(doc, remote, "Download URL available", item.downloadUrl.empty() ? "false" : "true");

    yyjson_mut_val* drive = AddPropertiesSection(doc, sections, "Drive");
    AddPropertiesFieldWide(doc, drive, "Drive ID", context.driveId);
    AddPropertiesFieldWide(doc, drive, "Site ID", context.siteId);
    AddPropertiesFieldWide(doc, drive, "Drive name", context.driveDisplayName);
    AddPropertiesFieldWide(doc, drive, "Volume label", context.driveVolumeLabel);
    AddPropertiesFieldWide(doc, drive, "Drive web URL", context.driveWebUrl);

    yyjson_mut_val* connection = AddPropertiesSection(doc, sections, "Connection");
    AddPropertiesFieldWide(doc, connection, "Connection name", context.connectionName);
    AddPropertiesFieldWide(doc, connection, "User", context.profile.userName);
    AddPropertiesFieldWide(doc, connection, "Authentication", context.profile.authMode);
    AddPropertiesField(doc, connection, "Save password", context.profile.savePassword ? "true" : "false");
    AddPropertiesFieldWide(doc, connection, "Authority", context.authority);

    yyjson_mut_val* timestamps = AddPropertiesSection(doc, sections, "Timestamps");
    AddPropertiesField(doc, timestamps, "Created", FormatFileTimeLocal(item.creationTime));
    AddPropertiesField(doc, timestamps, "Modified", FormatFileTimeLocal(item.lastWriteTime));
    AddPropertiesField(doc, timestamps, "Accessed", FormatFileTimeLocal(item.lastAccessTime));
    AddPropertiesField(doc, timestamps, "Changed", FormatFileTimeLocal(item.changeTime));

    if (rawRoot)
    {
        yyjson_mut_val* graph = AddPropertiesSection(doc, sections, "Graph");
        AddJsonScalarFields(doc, graph, rawRoot, "", 0u);
    }

    return WriteMutableJsonDocument(doc, jsonUtf8Out);
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
    Common::Json::UniqueMutableDocument docOwner{doc};

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
    Common::Json::UniqueMutableDocument docOwner{doc};

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

[[nodiscard]] HRESULT ParseUploadUrl(std::string_view bodyUtf8, ValidatedPreauthenticatedUploadUrl& uploadUrlOut) noexcept
{
    SecureClear(uploadUrlOut.value);

    yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
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

    std::wstring rawUploadUrl = TryGetJsonString(root, "uploadUrl").value_or(L"");
    auto clearRawUploadUrl    = wil::scope_exit([&] { SecureClear(rawUploadUrl); });
    if (rawUploadUrl.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    return ValidatePreauthenticatedUploadUrl(rawUploadUrl, uploadUrlOut);
}

[[nodiscard]] HRESULT ParseNextExpectedUploadOffset(std::string_view bodyUtf8, uint64_t totalBytes, uint64_t& offsetOut) noexcept
{
    offsetOut       = 0u;
    yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
    if (! doc)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    Common::Json::UniqueDocument docOwner{doc};

    yyjson_val* root   = yyjson_doc_get_root(doc);
    yyjson_val* ranges = root && yyjson_is_obj(root) ? yyjson_obj_get(root, "nextExpectedRanges") : nullptr;
    if (! ranges || ! yyjson_is_arr(ranges) || yyjson_arr_size(ranges) == 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    bool foundOffset  = false;
    uint64_t minimum  = (std::numeric_limits<uint64_t>::max)();
    size_t index      = 0u;
    size_t max        = 0u;
    yyjson_val* value = nullptr;
    yyjson_arr_foreach(ranges, index, max, value)
    {
        if (! value || ! yyjson_is_str(value))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        const std::string_view range(yyjson_get_str(value), yyjson_get_len(value));
        const size_t dash = range.find('-');
        if (dash == std::string_view::npos || dash == 0u || range.find('-', dash + 1u) != std::string_view::npos)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        uint64_t start         = 0u;
        const auto startResult = std::from_chars(range.data(), range.data() + dash, start);
        if (startResult.ec != std::errc{} || startResult.ptr != range.data() + dash || start >= totalBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        if (dash + 1u < range.size())
        {
            uint64_t end         = 0u;
            const auto endResult = std::from_chars(range.data() + dash + 1u, range.data() + range.size(), end);
            if (endResult.ec != std::errc{} || endResult.ptr != range.data() + range.size() || end < start || end >= totalBytes)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }

        minimum     = (std::min)(minimum, start);
        foundOffset = true;
    }

    if (! foundOffset)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    offsetOut = minimum;
    return S_OK;
}

[[nodiscard]] HRESULT BuildDriveContext(FileSystemMicrosoftDrive& fs, std::wstring_view rawPath, DriveContext& contextOut) noexcept
{
#if defined(_DEBUG)
    if (g_debugMicrosoftDriveUseSyntheticContextForSelfTest.load(std::memory_order_acquire))
    {
        constexpr std::wstring_view kDebugPrefix = L"/@conn:microsoft-drive-selftest";
        const std::wstring canonicalPath         = CanonicalizeInputPath(rawPath);
        if (! OrdinalString::StartsWithNoCase(canonicalPath, kDebugPrefix))
        {
            return HRESULT_FROM_WIN32(ERROR_NOT_CONNECTED);
        }

        contextOut                     = {};
        contextOut.connectionName      = L"microsoft-drive-selftest";
        contextOut.profile.name        = contextOut.connectionName;
        contextOut.profile.pluginId    = L"builtin/file-system-onedrive-personal";
        contextOut.profile.authMode    = L"oauth2Pkce";
        contextOut.authority           = L"consumers";
        contextOut.scopeText           = L"offline_access Files.ReadWrite User.Read openid profile";
        contextOut.driveId             = L"drive-selftest";
        contextOut.driveDisplayName    = L"Microsoft Drive SelfTest";
        contextOut.driveVolumeLabel    = L"Microsoft Drive SelfTest";
        contextOut.persistRefreshToken = false;
        contextOut.drivePath           = canonicalPath.size() == kDebugPrefix.size() ? L"/" : canonicalPath.substr(kDebugPrefix.size());
        return S_OK;
    }
#endif

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
        hr = SendAuthenticatedGraphHttpRequest(fs.SnapshotSettings(),
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
        hr = SendAuthenticatedGraphHttpRequest(fs.SnapshotSettings(), L"GET", siteUrl, accessToken.c_str(), headers, nullptr, 0, nullptr, true, siteResponse);
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
        hr = SendAuthenticatedGraphHttpRequest(fs.SnapshotSettings(), L"GET", driveUrl, accessToken.c_str(), headers, nullptr, 0, nullptr, true, driveResponse);
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

        hr = SendAuthenticatedGraphHttpRequest(fs.SnapshotSettings(),
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
        LogHttpResponseFailure(L"GET", url, response.statusCode, response.requestId, response.body.size());
        return HresultFromGraphError(response.statusCode, response.body);
    }

    return ParseItemMetadata(response.body, itemOut);
}

[[nodiscard]] HRESULT ListDirectory(FileSystemMicrosoftDrive& fs,
                                    const DriveContext& context,
                                    std::wstring_view drivePath,
                                    std::vector<FilesInformationMicrosoftDrive::Entry>& entriesOut,
                                    bool* incompleteDueToInvalidChildNameOut    = nullptr,
                                    const std::function<HRESULT()>& checkCancel = {}) noexcept
{
    entriesOut.clear();
    if (incompleteDueToInvalidChildNameOut)
    {
        *incompleteDueToInvalidChildNameOut = false;
    }

    const FileSystemMicrosoftDrive::Settings settings = fs.SnapshotSettings();
    std::wstring nextUrl                              = BuildGraphChildrenUrl(context, drivePath, settings.pageSize);
    auto clearNextUrl                                 = wil::scope_exit([&] { SecureClear(nextUrl); });
    const uint64_t nowTickMs                          = GetTickCount64();
    const uint64_t pagingDurationMs                   = std::clamp<uint64_t>(static_cast<uint64_t>(settings.requestTimeoutMs) * 10ull, 60'000ull, 600'000ull);
    Common::Paging::Limits pagingLimits{
        .deadlineTickMs = Common::Paging::DeadlineFromNow(nowTickMs, pagingDurationMs),
    };
    if (checkCancel)
    {
        pagingLimits.cancellationProbe  = [](void* cookie) noexcept -> HRESULT { return (*static_cast<const std::function<HRESULT()>*>(cookie))(); };
        pagingLimits.cancellationCookie = const_cast<std::function<HRESULT()>*>(&checkCancel);
    }
    Common::Paging::WideContinuationGuard pager(pagingLimits);
    bool firstPage = true;
    while (! nextUrl.empty())
    {
        const HRESULT pageBoundaryHr = firstPage ? pager.BeginFirstPage(nextUrl, GetTickCount64()) : pager.BeginContinuation(nextUrl, GetTickCount64());
        firstPage                    = false;
        if (FAILED(pageBoundaryHr))
        {
            Debug::Warning(L"Microsoft Drive: rejected bounded Graph pagination. method='GET' target='{}' hr=0x{:08X}.",
                           DescribeHttpRequestTarget(nextUrl),
                           static_cast<unsigned long>(pageBoundaryHr));
            return pageBoundaryHr;
        }

        HttpResponse response{};
        const HRESULT hr = SendGraphJsonRequest(fs, context, L"GET", nextUrl, {}, {}, true, response);
        if (FAILED(hr))
        {
            return hr;
        }

        if (response.statusCode < 200u || response.statusCode >= 300u)
        {
            LogHttpResponseFailure(L"GET", nextUrl, response.statusCode, response.requestId, response.body.size());
            return HresultFromGraphError(response.statusCode, response.body);
        }

        std::vector<FilesInformationMicrosoftDrive::Entry> pageEntries;
        std::wstring nextLink;
        auto clearNextLink                       = wil::scope_exit([&] { SecureClear(nextLink); });
        bool pageIncompleteDueToInvalidChildName = false;
        const HRESULT parseHr                    = ParseChildren(response.body, pageEntries, nextLink, &pageIncompleteDueToInvalidChildName);
        if (FAILED(parseHr))
        {
            return parseHr;
        }
        if (pageIncompleteDueToInvalidChildName && incompleteDueToInvalidChildNameOut)
        {
            *incompleteDueToInvalidChildNameOut = true;
        }

        const HRESULT pageHr = pager.CompletePage(pageEntries.size(), response.body.size(), ! nextLink.empty(), nextLink, GetTickCount64());
        if (FAILED(pageHr))
        {
            Debug::Warning(L"Microsoft Drive: rejected invalid or oversized Graph page sequence. hr=0x{:08X}.", static_cast<unsigned long>(pageHr));
            return pageHr;
        }

        entriesOut.insert(entriesOut.end(), std::make_move_iterator(pageEntries.begin()), std::make_move_iterator(pageEntries.end()));
        SecureClear(nextUrl);
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

[[nodiscard]] bool IsNotFoundStatus(HRESULT hr) noexcept;

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

    const std::wstring createUrl = BuildGraphCreateDirectoryUrl(context, parent.id);
    for (int attempt = 0; attempt < 4; ++attempt)
    {
        HttpResponse response{};
        hr = SendGraphJsonRequest(fs, context, L"POST", createUrl, bodyUtf8, {}, false, response);
        if (FAILED(hr))
        {
            return hr;
        }

        if (response.statusCode == 409u)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }
        if (response.statusCode >= 200u && response.statusCode < 300u)
        {
            return S_OK;
        }

        if (ShouldRetryGraphHttpResponse(true, response.statusCode, attempt))
        {
            ItemMetadata created{};
            const HRESULT probeHr = GetItemMetadata(fs, context, drivePath, false, created);
            if (SUCCEEDED(probeHr))
            {
                return created.isFolder ? S_OK : HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            }
            if (! IsNotFoundStatus(probeHr))
            {
                return probeHr;
            }

            LogAndSleepBeforeGraphRetry(L"POST", createUrl, response.statusCode, response.retryAfter, attempt);
            continue;
        }

        return HresultFromGraphError(response.statusCode, response.body);
    }

    return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
}

[[nodiscard]] HRESULT DeleteItemById(FileSystemMicrosoftDrive& fs, const DriveContext& context, std::wstring_view itemId) noexcept
{
    HttpResponse response{};
    HRESULT hr = SendGraphJsonRequest(fs, context, L"DELETE", BuildGraphItemByIdUrl(context, itemId), {}, {}, true, response);
    if (FAILED(hr))
    {
        return hr;
    }

    return response.statusCode == 204u ? S_OK : HresultFromGraphError(response.statusCode, response.body);
}

[[nodiscard]] HRESULT DeleteItemByPath(FileSystemMicrosoftDrive& fs, const DriveContext& context, std::wstring_view drivePath) noexcept
{
    ItemMetadata item{};
    HRESULT hr = GetItemMetadata(fs, context, drivePath, false, item);
    if (FAILED(hr))
    {
        return hr;
    }

    return DeleteItemById(fs, context, item.id);
}

[[nodiscard]] HRESULT MoveItemById(FileSystemMicrosoftDrive& fs,
                                   const DriveContext& context,
                                   std::wstring_view itemId,
                                   std::wstring_view newName,
                                   std::wstring_view parentId,
                                   bool includeParent) noexcept
{
    std::string bodyUtf8;
    HRESULT hr = BuildJsonBodyForMoveRename(newName, parentId, includeParent, bodyUtf8);
    if (FAILED(hr))
    {
        return hr;
    }

    HttpResponse response{};
    hr = SendGraphJsonRequest(fs, context, L"PATCH", BuildGraphItemByIdUrl(context, itemId), bodyUtf8, {}, true, response);
    if (FAILED(hr))
    {
        return hr;
    }

    return (response.statusCode >= 200u && response.statusCode < 300u) ? S_OK : HresultFromGraphError(response.statusCode, response.body);
}

[[nodiscard]] HRESULT BuildRollbackLeafName(std::wstring& leafOut) noexcept
{
    return Common::Paths::BuildUniqueSiblingName(
        std::wstring_view{}, std::wstring_view(L".redsalamander-rollback-"), std::wstring_view{}, (std::numeric_limits<size_t>::max)(), leafOut);
}

[[nodiscard]] HRESULT PrepareOverwriteBackup(FileSystemMicrosoftDrive& fs,
                                             const DriveContext& context,
                                             const ItemMetadata& existingDestination,
                                             std::wstring_view destinationParentPath,
                                             std::wstring& backupItemIdOut) noexcept
{
    backupItemIdOut.clear();

    if (existingDestination.isFolder)
    {
        return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
    }

    for (unsigned int attempt = 0; attempt < 16u; ++attempt)
    {
        std::wstring backupLeaf;
        HRESULT hr = BuildRollbackLeafName(backupLeaf);
        if (FAILED(hr))
        {
            return hr;
        }
        if (attempt != 0u)
        {
            backupLeaf.append(std::format(L"-{}", attempt));
        }

        const std::wstring backupPath = JoinPath(destinationParentPath, backupLeaf);
        ItemMetadata existingBackup{};
        hr = GetItemMetadata(fs, context, backupPath, false, existingBackup);
        if (SUCCEEDED(hr))
        {
            continue;
        }
        if (! IsNotFoundStatus(hr))
        {
            return hr;
        }

        hr = MoveItemById(fs, context, existingDestination.id, backupLeaf, {}, false);
        if (FAILED(hr))
        {
            return hr;
        }

        backupItemIdOut = existingDestination.id;
        return S_OK;
    }

    return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
}

// Per-child conflict callback: (sourceDisplayPath, destinationDisplayPath, status, action out).
using MoveIssueReporter = std::function<HRESULT(const wchar_t*, const wchar_t*, HRESULT, FileSystemIssueAction&)>;
using CancelProbe       = std::function<HRESULT()>;

struct MoveCommitResult
{
    bool primaryMutationCommitted = false;
    HRESULT cleanupStatus         = S_OK;
    HRESULT rollbackStatus        = S_OK;
};

inline constexpr unsigned int kMicrosoftDriveMergeMaxDepth = 64u;

[[nodiscard]] HRESULT MoveOrRenameItem(FileSystemMicrosoftDrive& fs,
                                       const DriveContext& sourceContext,
                                       const DriveContext& destinationContext,
                                       std::wstring_view sourcePath,
                                       std::wstring_view destinationPath,
                                       FileSystemFlags flags,
                                       const MoveIssueReporter& reportIssue = {},
                                       bool allowDirectoryMerge             = true,
                                       const CancelProbe& checkCancel       = {},
                                       MoveCommitResult* commitResultOut    = nullptr) noexcept;

[[nodiscard]] bool IsMergeChildPartialFailure(HRESULT hr) noexcept
{
    return hr == HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
}

// Moves every child of sourceFolderPath into the existing destinationFolderPath (server-side
// moves), recursing where both sides have a folder of the same name. Collisions are resolved
// per child via reportIssue. The source folder is always retained after a successful merge
// because Graph has no atomic "delete only if still empty" operation; this prevents deleting
// a child added concurrently after the merge listing.
[[nodiscard]] HRESULT MergeMoveFolderIntoExisting(FileSystemMicrosoftDrive& fs,
                                                  const DriveContext& sourceContext,
                                                  const DriveContext& destinationContext,
                                                  const std::wstring& sourceFolderPath,
                                                  const std::wstring& destinationFolderPath,
                                                  FileSystemFlags flags,
                                                  const MoveIssueReporter& reportIssue,
                                                  const CancelProbe& checkCancel,
                                                  bool& anySkipped,
                                                  bool& subtreeFullyMovedOut,
                                                  unsigned int depth = 0) noexcept
{
    subtreeFullyMovedOut = false;

    if (depth >= kMicrosoftDriveMergeMaxDepth)
    {
        return HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW);
    }

    if (checkCancel)
    {
        const HRESULT cancelHr = checkCancel();
        if (FAILED(cancelHr))
        {
            return cancelHr;
        }
    }

    std::vector<FilesInformationMicrosoftDrive::Entry> children;
    bool sourceEnumerationIncomplete = false;
    HRESULT hr                       = ListDirectory(fs, sourceContext, sourceFolderPath, children, &sourceEnumerationIncomplete, checkCancel);
    if (FAILED(hr))
    {
        return hr;
    }

    const bool allowOverwriteFlag = (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
    bool allChildrenMoved         = ! sourceEnumerationIncomplete;
    if (sourceEnumerationIncomplete)
    {
        anySkipped = true;
    }

    for (const auto& child : children)
    {
        if (checkCancel)
        {
            hr = checkCancel();
            if (FAILED(hr))
            {
                return hr;
            }
        }

        const std::wstring childSource      = JoinPath(sourceFolderPath, child.name);
        const std::wstring childDestination = JoinPath(destinationFolderPath, child.name);
        const bool childIsFolder            = (child.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;

        ItemMetadata existingChild{};
        const HRESULT probeHr = GetItemMetadata(fs, destinationContext, childDestination, false, existingChild);
        if (FAILED(probeHr))
        {
            if (! IsNotFoundStatus(probeHr))
            {
                return probeHr;
            }

            hr = MoveOrRenameItem(
                fs, sourceContext, destinationContext, childSource, childDestination, static_cast<FileSystemFlags>(flags & ~FILESYSTEM_FLAG_ALLOW_OVERWRITE));
            if (FAILED(hr))
            {
                if (IsMergeChildPartialFailure(hr))
                {
                    anySkipped       = true;
                    allChildrenMoved = false;
                    continue;
                }
                return hr;
            }
            continue;
        }

        if (childIsFolder && existingChild.isFolder)
        {
            bool childSubtreeFullyMoved = false;
            hr                          = MergeMoveFolderIntoExisting(fs,
                                                                      sourceContext,
                                                                      destinationContext,
                                                                      childSource,
                                                                      childDestination,
                                                                      flags,
                                                                      reportIssue,
                                                                      checkCancel,
                                                                      anySkipped,
                                                                      childSubtreeFullyMoved,
                                                                      depth + 1u);
            if (FAILED(hr))
            {
                if (IsMergeChildPartialFailure(hr))
                {
                    anySkipped       = true;
                    allChildrenMoved = false;
                    continue;
                }
                return hr;
            }
            if (! childSubtreeFullyMoved)
            {
                allChildrenMoved = false;
            }
            continue;
        }

        const bool typeMismatch = childIsFolder != existingChild.isFolder;
        bool overwriteChild     = allowOverwriteFlag && ! typeMismatch;
        bool childSkipped       = false;
        if (! overwriteChild)
        {
            if (! reportIssue)
            {
                return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
            }
            else
            {
                while (! childSkipped && ! overwriteChild)
                {
                    FileSystemIssueAction action = FileSystemIssueAction::Cancel;
                    hr                           = reportIssue(childSource.c_str(), childDestination.c_str(), HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS), action);
                    if (FAILED(hr))
                    {
                        return hr;
                    }

                    switch (action)
                    {
                        case FileSystemIssueAction::Overwrite:
                        case FileSystemIssueAction::ReplaceReadOnly:
                            if (typeMismatch)
                            {
                                // Graph drives cannot replace across item kinds; unresolvable here.
                                childSkipped = true;
                                break;
                            }
                            overwriteChild = true;
                            break;
                        case FileSystemIssueAction::Retry:
                        {
                            ItemMetadata reprobe{};
                            const HRESULT retryHr = GetItemMetadata(fs, destinationContext, childDestination, false, reprobe);
                            if (SUCCEEDED(retryHr))
                            {
                                continue;
                            }
                            if (! IsNotFoundStatus(retryHr))
                            {
                                return retryHr;
                            }
                            overwriteChild = true; // conflict resolved externally; plain move below
                            break;
                        }
                        case FileSystemIssueAction::Skip: childSkipped = true; break;
                        case FileSystemIssueAction::PermanentDelete:
                        case FileSystemIssueAction::Cancel:
                        case FileSystemIssueAction::None:
                        default: return HRESULT_FROM_WIN32(ERROR_CANCELLED);
                    }
                }
            }
        }

        if (childSkipped)
        {
            anySkipped       = true;
            allChildrenMoved = false;
            continue;
        }

        hr = MoveOrRenameItem(
            fs, sourceContext, destinationContext, childSource, childDestination, static_cast<FileSystemFlags>(flags | FILESYSTEM_FLAG_ALLOW_OVERWRITE));
        if (FAILED(hr))
        {
            if (IsMergeChildPartialFailure(hr))
            {
                anySkipped       = true;
                allChildrenMoved = false;
                continue;
            }
            return hr;
        }
    }

    if (! allChildrenMoved)
    {
        return S_OK;
    }

    if (checkCancel)
    {
        hr = checkCancel();
        if (FAILED(hr))
        {
            return hr;
        }
    }

    // Graph item DELETE is recursive and this provider has no atomic "delete only if still empty"
    // primitive. Even an immediately preceding empty re-list leaves a race in which another client can
    // add a child before DELETE. Keep the drained source folder and report partial cleanup; this is less
    // convenient than a perfect folder MOVE, but it cannot erase a concurrently-added child's only copy.
    anySkipped = true;
    return S_OK;
}

[[nodiscard]] HRESULT MoveOrRenameItem(FileSystemMicrosoftDrive& fs,
                                       const DriveContext& sourceContext,
                                       const DriveContext& destinationContext,
                                       std::wstring_view sourcePath,
                                       std::wstring_view destinationPath,
                                       FileSystemFlags flags,
                                       const MoveIssueReporter& reportIssue,
                                       bool allowDirectoryMerge,
                                       const CancelProbe& checkCancel,
                                       MoveCommitResult* commitResultOut) noexcept
{
    if (commitResultOut)
    {
        *commitResultOut = {};
    }
    if (! OrdinalString::EqualsNoCase(sourceContext.connectionName, destinationContext.connectionName) ||
        ! OrdinalString::EqualsNoCase(sourceContext.driveId, destinationContext.driveId))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    const std::wstring normalizedSource      = TrimTrailingSlashPreserveRoot(NormalizePluginPath(sourcePath));
    const std::wstring normalizedDestination = TrimTrailingSlashPreserveRoot(NormalizePluginPath(destinationPath));
    const bool samePathIdentity              = OrdinalString::EqualsNoCase(normalizedSource, normalizedDestination);
    if (! samePathIdentity && IsSameOrDescendantPath(normalizedSource, normalizedDestination))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER);
    }

    ItemMetadata sourceItem{};
    HRESULT hr = GetItemMetadata(fs, sourceContext, normalizedSource, false, sourceItem);
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring sourceParentPath;
    std::wstring sourceLeafName;
    hr = SplitParentAndLeaf(normalizedSource, sourceParentPath, sourceLeafName);
    if (FAILED(hr))
    {
        return hr;
    }

    ItemMetadata sourceParent{};
    hr = EnsureParentMetadata(fs, sourceContext, sourceParentPath, sourceParent);
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring destinationParentPath;
    std::wstring destinationName;
    hr = SplitParentAndLeaf(normalizedDestination, destinationParentPath, destinationName);
    if (FAILED(hr))
    {
        return hr;
    }

    std::wstring backupItemId;
    ItemMetadata existingDestination{};
    hr = GetItemMetadata(fs, destinationContext, normalizedDestination, false, existingDestination);
    if (SUCCEEDED(hr))
    {
        const bool destinationIsSource = OrdinalString::EqualsNoCase(existingDestination.id, sourceItem.id);
        if (destinationIsSource && normalizedSource == normalizedDestination)
        {
            return S_OK;
        }

        if (! destinationIsSource && allowDirectoryMerge && sourceItem.isFolder && existingDestination.isFolder)
        {
            // Directory-onto-directory MERGES (normative rule): folder existence is never an
            // overwrite conflict; only per-child collisions prompt.
            bool anySkipped        = false;
            bool subtreeFullyMoved = false;
            hr                     = MergeMoveFolderIntoExisting(
                fs, sourceContext, destinationContext, normalizedSource, normalizedDestination, flags, reportIssue, checkCancel, anySkipped, subtreeFullyMoved);
            if (FAILED(hr))
            {
                return hr;
            }
            return anySkipped ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : S_OK;
        }

        if (! destinationIsSource && (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) == 0)
        {
            return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS);
        }

        if (! destinationIsSource)
        {
            hr = PrepareOverwriteBackup(fs, destinationContext, existingDestination, destinationParentPath, backupItemId);
            if (FAILED(hr))
            {
                return hr;
            }
        }
    }
    else if (! IsNotFoundStatus(hr))
    {
        return hr;
    }

    ItemMetadata destinationParent{};
    hr = EnsureParentMetadata(fs, destinationContext, destinationParentPath, destinationParent);
    if (FAILED(hr))
    {
        if (! backupItemId.empty())
        {
            const HRESULT restoreHr = MoveItemById(fs, destinationContext, backupItemId, destinationName, {}, false);
            if (commitResultOut)
            {
                commitResultOut->rollbackStatus = restoreHr;
            }
            return FAILED(restoreHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : hr;
        }
        return hr;
    }

    const bool includeParent = ! OrdinalString::EqualsNoCase(destinationParent.id, sourceParent.id);
    if (checkCancel)
    {
        hr = checkCancel();
        if (FAILED(hr))
        {
            if (! backupItemId.empty())
            {
                const HRESULT restoreHr = MoveItemById(fs, destinationContext, backupItemId, destinationName, {}, false);
                if (commitResultOut)
                {
                    commitResultOut->rollbackStatus = restoreHr;
                }
                return FAILED(restoreHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : hr;
            }
            return hr;
        }
    }
    hr = MoveItemById(fs, sourceContext, sourceItem.id, destinationName, destinationParent.id, includeParent);
    if (FAILED(hr))
    {
        if (! backupItemId.empty())
        {
            const HRESULT restoreHr = MoveItemById(fs, destinationContext, backupItemId, destinationName, {}, false);
            if (commitResultOut)
            {
                commitResultOut->rollbackStatus = restoreHr;
            }
            return FAILED(restoreHr) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : hr;
        }
        return hr;
    }

    if (commitResultOut)
    {
        commitResultOut->primaryMutationCommitted = true;
    }

    if (! backupItemId.empty())
    {
        const HRESULT cleanupHr = DeleteItemById(fs, destinationContext, backupItemId);
        if (FAILED(cleanupHr))
        {
            if (commitResultOut)
            {
                commitResultOut->cleanupStatus = cleanupHr;
            }
            Debug::Warning(L"Microsoft Drive: move committed but overwrite-backup cleanup remains pending. hr=0x{:08X}.",
                           static_cast<unsigned long>(cleanupHr));
            return S_OK;
        }
    }

    return S_OK;
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
    MicrosoftDriveRangedFileReader(FileSystemMicrosoftDrive& fileSystem,
                                   DriveContext context,
                                   std::wstring drivePath,
                                   uint64_t sizeBytes,
                                   std::wstring downloadUrl,
                                   std::wstring eTag) noexcept
        : _fileSystem(&fileSystem),
          _context(std::move(context)),
          _drivePath(std::move(drivePath)),
          _sizeBytes(sizeBytes),
          _downloadUrl(std::move(downloadUrl)),
          _eTag(std::move(eTag)),
          _settings(fileSystem.SnapshotSettings())
    {
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
    ~MicrosoftDriveRangedFileReader() = default;

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

        if (! _eTag.empty() && ! item.eTag.empty() && _eTag != item.eTag)
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
        if (item.sizeBytes != _sizeBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
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
        chunkBytes          = std::min<uint64_t>(chunkBytes, 16ull * 1024ull * 1024ull);
        chunkBytes          = std::min<uint64_t>(chunkBytes, _sizeBytes - offset);

        for (int attempt = 0; attempt < 2; ++attempt)
        {
            HRESULT hr = _downloadUrl.empty() ? RefreshDownloadUrl() : S_OK;
            if (FAILED(hr))
            {
                return hr;
            }

            const uint64_t endOffset                = offset + chunkBytes - 1ull;
            const std::array<HttpHeader, 3> headers = {
                HttpHeader{L"Accept", L"application/octet-stream"},
                HttpHeader{L"Range", std::format(L"bytes={}-{}", offset, endOffset)},
                HttpHeader{L"If-Match", _eTag},
            };
            const std::span<const HttpHeader> requestHeaders(headers.data(), _eTag.empty() ? 2u : headers.size());

            HttpResponse response{};
            hr = SendHttpRequest(_settings, L"GET", _downloadUrl, nullptr, requestHeaders, nullptr, 0, nullptr, true, false, response);
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
    wil::com_ptr<FileSystemMicrosoftDrive> _fileSystem;
    DriveContext _context;
    std::wstring _drivePath;
    uint64_t _position    = 0;
    uint64_t _sizeBytes   = 0;
    uint64_t _cacheOffset = 0;
    std::wstring _downloadUrl;
    std::wstring _eTag;
    FileSystemMicrosoftDrive::Settings _settings;
    std::vector<std::byte> _cache;
};

class MicrosoftDriveFileWriter final : public IFileWriter, public IFileWriterExpectedSize
{
public:
    MicrosoftDriveFileWriter(FileSystemMicrosoftDrive& fileSystem, std::wstring destinationPath, FileSystemFlags flags, wil::unique_hfile tempFile) noexcept
        : _fileSystem(&fileSystem),
          _destinationPath(std::move(destinationPath)),
          _flags(flags),
          _tempFile(std::move(tempFile))
    {
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
        if (riid == __uuidof(IFileWriterExpectedSize))
        {
            *ppvObject = static_cast<IFileWriterExpectedSize*>(this);
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

    HRESULT STDMETHODCALLTYPE SetExpectedSize(uint64_t sizeBytes) noexcept override
    {
        if (_committed || _position != 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        _expectedSize    = sizeBytes;
        _hasExpectedSize = true;
        if (sizeBytes != 0u)
        {
            // A known non-empty size lets Write stream through a Graph upload session. Closing
            // the delete-on-close fallback here keeps temporary-disk use at zero for bridge copies.
            _tempFile.reset();
            _streaming = true;
        }
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
        if (_streaming)
        {
            return WriteStreaming(buffer, bytesToWrite, bytesWritten);
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
        if (! _streaming || ! _streamInitialized || _committed || _uploadUrl.empty() || _lastUploadStatus == 200u || _lastUploadStatus == 201u)
        {
            return;
        }

        HttpResponse response{};
        const HRESULT hr = SendPreauthenticatedUploadRequest(_streamSettings, L"DELETE", _uploadUrl, {}, nullptr, 0u, false, response);
        if (FAILED(hr) || (response.statusCode < 200u || response.statusCode >= 300u) && response.statusCode != 404u)
        {
            Debug::Warning(L"Microsoft Drive: upload-session cleanup failed. destination='{}' hr=0x{:08X} status={}.",
                           _destinationPath,
                           static_cast<unsigned long>(hr),
                           response.statusCode);
        }
    }

    [[nodiscard]] HRESULT InitializeStreamingUpload() noexcept;
    [[nodiscard]] HRESULT FlushStreamingChunk() noexcept;
    [[nodiscard]] HRESULT WriteStreaming(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept;
    [[nodiscard]] HRESULT UploadSimple(const DriveContext& context, std::wstring_view drivePath, uint64_t fileSize, bool allowOverwrite) noexcept;
    [[nodiscard]] HRESULT UploadWithSession(const DriveContext& context, std::wstring_view drivePath, uint64_t fileSize, bool allowOverwrite) noexcept;

    std::atomic_ulong _refCount{1};
    wil::com_ptr<FileSystemMicrosoftDrive> _fileSystem;
    std::wstring _destinationPath;
    FileSystemFlags _flags = FILESYSTEM_FLAG_NONE;
    wil::unique_hfile _tempFile;
    uint64_t _position      = 0;
    uint64_t _expectedSize  = 0;
    uint64_t _uploadedBytes = 0;
    FileSystemMicrosoftDrive::Settings _streamSettings;
    ValidatedPreauthenticatedUploadUrl _uploadUrl;
    std::vector<std::byte> _uploadBuffer;
    size_t _uploadChunkBytes = 0;
    DWORD _lastUploadStatus  = 0;
    bool _hasExpectedSize    = false;
    bool _streaming          = false;
    bool _streamInitialized  = false;
    bool _committed          = false;
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

    _configurationJsonStorage[0] = "{}";
    _configurationJsonStorage[1] = "{}";
    _propertiesJson              = "{}";
    _driveFileSystem             = _metaData.shortId ? _metaData.shortId : L"";

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

    *schemaJsonUtf8 = StaticConfigurationSchema(_mode);
    return S_OK;
}

const char* GetFileSystemMicrosoftDriveStaticConfigurationSchema(FileSystemMicrosoftDriveMode mode) noexcept
{
    return FileSystemMicrosoftDrive::StaticConfigurationSchema(mode);
}

const char* FileSystemMicrosoftDrive::StaticConfigurationSchema(FileSystemMicrosoftDriveMode mode) noexcept
{
    static_cast<void>(mode);
    return kSchemaJson;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    Settings nextSettings{};
    std::string nextConfiguration = "{}";
    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        nextConfiguration                   = configurationJsonUtf8;
        Common::Json::ObjectDocument parsed = Common::Json::ParseObjectDocument(nextConfiguration, YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (! parsed)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        yyjson_val* root = parsed.root;
        if (const auto value = TryGetJsonString(root, "clientId"); value.has_value())
        {
            nextSettings.clientId = value.value();
        }
        if (const auto value = TryGetJsonUInt(root, "connectTimeoutMs"); value.has_value() && value.value() >= 1u)
        {
            nextSettings.connectTimeoutMs = static_cast<uint32_t>((std::min)(value.value(), 600'000ull));
        }
        if (const auto value = TryGetJsonUInt(root, "requestTimeoutMs"); value.has_value() && value.value() >= 1u)
        {
            nextSettings.requestTimeoutMs = static_cast<uint32_t>((std::min)(value.value(), 600'000ull));
        }
        if (const auto value = TryGetJsonUInt(root, "pageSize"); value.has_value() && value.value() >= 1u)
        {
            nextSettings.pageSize = static_cast<uint32_t>((std::min)(value.value(), 999ull));
        }
        if (const auto value = TryGetJsonUInt(root, "uploadChunkMiB"); value.has_value() && value.value() >= 1u)
        {
            nextSettings.uploadChunkMiB = static_cast<uint32_t>((std::min)(value.value(), 32ull));
        }
    }

    std::lock_guard lock(_stateMutex);
    for (auto& [key, token] : _tokenCacheByConnectionName)
    {
        SecureClear(token.accessToken);
    }
    _tokenCacheByConnectionName.clear();
    _driveCacheByConnectionName.clear();
    _settings = std::move(nextSettings);

    const size_t nextIndex               = 1u - _configurationJsonIndex;
    _configurationJsonStorage[nextIndex] = std::move(nextConfiguration);
    _configurationJsonIndex              = nextIndex;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (! configurationJsonUtf8)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    *configurationJsonUtf8 = _configurationJsonStorage[_configurationJsonIndex].c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (! pSomethingToSave)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    const auto& config = _configurationJsonStorage[_configurationJsonIndex];
    *pSomethingToSave  = (! config.empty() && config != "{}") ? TRUE : FALSE;
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
    _navigationMenuCallbackState.Set(callback, cookie);
    return S_OK;
}

bool FileSystemMicrosoftDrive::TryCaptureNavigationMenuCallback(NavigationMenuCallbackSnapshot& snapshot) noexcept
{
    return _navigationMenuCallbackState.TryCapture(snapshot);
}

HRESULT FileSystemMicrosoftDrive::InvokeNavigationMenuCallback(const NavigationMenuCallbackSnapshot& snapshot, const wchar_t* path) noexcept
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

namespace
{
// Shared core for MoveItem/MoveItems: drive contexts, caller-namespace conflict reporting and
// the merge-capable MoveOrRenameItem call. Item-result reporting stays with the caller so batch
// moves don't double-report per item.
HRESULT MoveSingleItemWithConflicts(FileSystemMicrosoftDrive& fs,
                                    const wchar_t* sourcePath,
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
    if (options != nullptr && options->sizeBytes != sizeof(FileSystemOptions))
    {
        return E_INVALIDARG;
    }

    DriveContext sourceContext{};
    HRESULT hr = BuildDriveContext(fs, sourcePath, sourceContext);
    if (FAILED(hr))
    {
        return hr;
    }

    DriveContext destinationContext{};
    hr = BuildDriveContext(fs, destinationPath, destinationContext);
    if (FAILED(hr))
    {
        return hr;
    }

    FileSystemOptions optionsState{};
    if (options != nullptr)
    {
        optionsState = *options;
    }
    optionsState.sizeBytes             = sizeof(FileSystemOptions);
    FileSystemOptions* callbackOptions = callback ? &optionsState : nullptr;

    const std::wstring sourceDriveRoot      = TrimTrailingSlashPreserveRoot(NormalizePluginPath(sourceContext.drivePath));
    const std::wstring destinationDriveRoot = TrimTrailingSlashPreserveRoot(NormalizePluginPath(destinationContext.drivePath));

    // Conflict prompts must name paths in the caller's namespace; merge children arrive in
    // drive-path form, so the suffix beyond the drive root is grafted onto the original path.
    const auto mapChildToOriginal =
        [](const std::wstring& originalRoot, const std::wstring& driveRoot, std::wstring_view childDrivePath) noexcept -> std::wstring
    {
        std::wstring display = originalRoot;
        if (childDrivePath.size() > driveRoot.size())
        {
            const std::wstring childPrefix(childDrivePath.substr(0, driveRoot.size()));
            if (OrdinalString::EqualsNoCase(childPrefix, driveRoot))
            {
                std::wstring_view suffix = childDrivePath.substr(driveRoot.size());
                while (! suffix.empty() && (suffix.front() == L'/' || suffix.front() == L'\\'))
                {
                    suffix.remove_prefix(1);
                }
                if (! suffix.empty())
                {
                    if (! display.empty() && display.back() != L'/' && display.back() != L'\\')
                    {
                        display.push_back(L'/');
                    }
                    display.append(suffix);
                }
            }
        }
        return display;
    };

    const std::wstring originalSource(sourcePath);
    const std::wstring originalDestination(destinationPath);

    const auto reportIssue =
        [&](const wchar_t* conflictSource, const wchar_t* conflictDestination, HRESULT status, FileSystemIssueAction& action) noexcept -> HRESULT
    {
        action = FileSystemIssueAction::Cancel;
        if (! callback)
        {
            return HRESULT_FROM_WIN32(ERROR_CANCELLED);
        }

        const std::wstring mappedSource      = mapChildToOriginal(originalSource, sourceDriveRoot, conflictSource ? conflictSource : L"");
        const std::wstring mappedDestination = mapChildToOriginal(originalDestination, destinationDriveRoot, conflictDestination ? conflictDestination : L"");
        const HRESULT issueHr =
            callback->FileSystemIssue(FILESYSTEM_MOVE, mappedSource.c_str(), mappedDestination.c_str(), status, &action, callbackOptions, cookie);
        return issueHr == E_ABORT ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : issueHr;
    };

    const CancelProbe checkCancel = callback ? CancelProbe(
                                                   [callback, cookie]() noexcept -> HRESULT
    {
        bool cancelled         = false;
        const HRESULT cancelHr = CheckShouldCancel(callback, cookie, cancelled);
        if (FAILED(cancelHr))
        {
            return cancelHr;
        }
        return cancelled ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
    })
                                             : CancelProbe{};

    return MoveOrRenameItem(fs,
                            sourceContext,
                            destinationContext,
                            sourceContext.drivePath,
                            destinationContext.drivePath,
                            flags,
                            callback ? MoveIssueReporter(reportIssue) : MoveIssueReporter{},
                            true,
                            checkCancel);
}
} // namespace

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::MoveItem(const wchar_t* sourcePath,
                                                             const wchar_t* destinationPath,
                                                             FileSystemFlags flags,
                                                             const FileSystemOptions* options,
                                                             IFileSystemCallback* callback,
                                                             void* cookie) noexcept
{
    const HRESULT hr         = MoveSingleItemWithConflicts(*this, sourcePath, destinationPath, flags, options, callback, cookie);
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

    // Rename never merges: renaming onto an existing folder stays an explicit conflict.
    hr = MoveOrRenameItem(*this, sourceContext, destinationContext, sourceContext.drivePath, destinationContext.drivePath, flags, MoveIssueReporter{}, false);
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
    const bool continueOnError         = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
    HRESULT firstFailure               = S_OK;
    bool hadFailure                    = false;
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

        if (! sourcePaths[i] || sourcePaths[i][0] == L'\0')
        {
            hr                       = E_INVALIDARG;
            const HRESULT callbackHr = ReportItemResult(callback, FILESYSTEM_COPY, count, i + 1u, i, sourcePaths[i], destinationFolder, hr, options, cookie);
            if (FAILED(callbackHr))
            {
                return callbackHr;
            }
            hadFailure = true;
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = hr;
            }
            if (! continueOnError)
            {
                return hr;
            }
            continue;
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
            hadFailure   = true;
            if (! continueOnError)
            {
                return hr;
            }
        }
    }

    return (hadFailure && continueOnError) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : firstFailure;
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
    const bool continueOnError         = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
    HRESULT firstFailure               = S_OK;
    bool hadFailure                    = false;
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

        if (! sourcePaths[i] || sourcePaths[i][0] == L'\0')
        {
            hr                       = E_INVALIDARG;
            const HRESULT callbackHr = ReportItemResult(callback, FILESYSTEM_MOVE, count, i + 1u, i, sourcePaths[i], destinationFolder, hr, options, cookie);
            if (FAILED(callbackHr))
            {
                return callbackHr;
            }
            hadFailure = true;
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = hr;
            }
            if (! continueOnError)
            {
                return hr;
            }
            continue;
        }

        std::wstring parentPath;
        std::wstring leafName;
        hr = SplitParentAndLeaf(CanonicalizeInputPath(sourcePaths[i]), parentPath, leafName);
        if (SUCCEEDED(hr))
        {
            const std::wstring destinationPath = JoinPath(destinationRoot, leafName);
            // Pass the real callback so directory merges can prompt per child; the batch-level
            // ReportItemResult below stays the single item-result reporter.
            hr = MoveSingleItemWithConflicts(*this, sourcePaths[i], destinationPath.c_str(), flags, options, callback, cookie);
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
            hadFailure   = true;
            if (! continueOnError)
            {
                return hr;
            }
        }
    }

    return (hadFailure && continueOnError) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : firstFailure;
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

    const bool continueOnError = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
    HRESULT firstFailure       = S_OK;
    bool hadFailure            = false;
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

        if (! paths[i] || paths[i][0] == L'\0')
        {
            hr                       = E_INVALIDARG;
            const HRESULT callbackHr = ReportItemResult(callback, FILESYSTEM_DELETE, count, i + 1u, i, paths[i], nullptr, hr, options, cookie);
            if (FAILED(callbackHr))
            {
                return callbackHr;
            }
            hadFailure = true;
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = hr;
            }
            if (! continueOnError)
            {
                return hr;
            }
            continue;
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
            hadFailure   = true;
            if (! continueOnError)
            {
                return hr;
            }
        }
    }

    return (hadFailure && continueOnError) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : firstFailure;
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

    for (unsigned long i = 0; i < count; ++i)
    {
        if (items[i].sizeBytes != sizeof(FileSystemRenamePair))
        {
            return E_INVALIDARG;
        }
    }

    const bool continueOnError = (flags & FILESYSTEM_FLAG_CONTINUE_ON_ERROR) != 0;
    HRESULT firstFailure       = S_OK;
    bool hadFailure            = false;
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

        if (! items[i].sourcePath || items[i].sourcePath[0] == L'\0' || ! items[i].newName || items[i].newName[0] == L'\0')
        {
            hr                       = E_INVALIDARG;
            const HRESULT callbackHr = ReportItemResult(callback, FILESYSTEM_RENAME, count, i + 1u, i, items[i].sourcePath, nullptr, hr, options, cookie);
            if (FAILED(callbackHr))
            {
                return callbackHr;
            }
            hadFailure = true;
            if (SUCCEEDED(firstFailure))
            {
                firstFailure = hr;
            }
            if (! continueOnError)
            {
                return hr;
            }
            continue;
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
            hadFailure   = true;
            if (! continueOnError)
            {
                return hr;
            }
        }
    }

    return (hadFailure && continueOnError) ? HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) : firstFailure;
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

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetTransferHints([[maybe_unused]] const wchar_t* path,
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

HRESULT STDMETHODCALLTYPE FileSystemMicrosoftDrive::GetStorageCharacteristics([[maybe_unused]] const wchar_t* path,
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
    auto* impl                   = new (std::nothrow)
        MicrosoftDriveRangedFileReader(*this, std::move(context), drivePath, item.sizeBytes, std::move(item.downloadUrl), std::move(item.eTag));
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

    std::string propertiesJson;
    hr = BuildItemPropertiesJson(context, item, propertiesJson);
    if (FAILED(hr))
    {
        return hr;
    }

    hr = SetPropertiesJson(std::move(propertiesJson));
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

HRESULT MicrosoftDriveFileWriter::InitializeStreamingUpload() noexcept
{
    if (_streamInitialized)
    {
        return S_OK;
    }
    if (! _hasExpectedSize || _expectedSize == 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
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

    std::string bodyUtf8;
    hr = BuildJsonBodyForUploadSession(allowOverwrite, bodyUtf8);
    if (FAILED(hr))
    {
        return hr;
    }

    HttpResponse sessionResponse{};
    auto clearSessionResponseBody = wil::scope_exit([&] { SecureClear(sessionResponse.body); });
    hr = SendGraphJsonRequest(*_fileSystem, context, L"POST", BuildGraphCreateUploadSessionUrl(context, normalizedPath), bodyUtf8, {}, false, sessionResponse);
    if (FAILED(hr))
    {
        return hr;
    }
    if (sessionResponse.statusCode < 200u || sessionResponse.statusCode >= 300u)
    {
        return HresultFromGraphError(sessionResponse.statusCode, sessionResponse.body);
    }

    hr = ParseUploadUrl(sessionResponse.body, _uploadUrl);
    if (FAILED(hr))
    {
        return hr;
    }

    _streamSettings             = _fileSystem->SnapshotSettings();
    const uint64_t chunkBytes64 = ComputeUploadChunkSizeBytes(_streamSettings);
    if (chunkBytes64 == 0u || chunkBytes64 > static_cast<uint64_t>((std::numeric_limits<size_t>::max)()) ||
        chunkBytes64 > static_cast<uint64_t>((std::numeric_limits<unsigned long>::max)()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    _uploadChunkBytes = static_cast<size_t>(chunkBytes64);
    _uploadBuffer.reserve(_uploadChunkBytes);
    _streamInitialized = true;
    return S_OK;
}

HRESULT MicrosoftDriveFileWriter::FlushStreamingChunk() noexcept
{
    unsigned int partialAcknowledgements = 0u;
    while (! _uploadBuffer.empty())
    {
        if (! _streamInitialized || _uploadedBytes > _expectedSize || _uploadBuffer.size() > (_expectedSize - _uploadedBytes))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const uint64_t chunkSize                = static_cast<uint64_t>(_uploadBuffer.size());
        const uint64_t endOffset                = _uploadedBytes + chunkSize - 1u;
        const std::array<HttpHeader, 3> headers = {
            HttpHeader{L"Accept", L"application/json"},
            HttpHeader{L"Content-Length", std::format(L"{}", chunkSize)},
            HttpHeader{L"Content-Range", std::format(L"bytes {}-{}/{}", _uploadedBytes, endOffset, _expectedSize)},
        };

        HttpResponse uploadResponse{};
        HRESULT hr =
            SendPreauthenticatedUploadRequest(_streamSettings, L"PUT", _uploadUrl, headers, _uploadBuffer.data(), _uploadBuffer.size(), true, uploadResponse);
        if (FAILED(hr))
        {
            return hr;
        }

        const bool isFinal = endOffset + 1u == _expectedSize;
        if (uploadResponse.statusCode == 200u || uploadResponse.statusCode == 201u)
        {
            if (! isFinal)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
            _lastUploadStatus = uploadResponse.statusCode;
            _uploadedBytes += chunkSize;
            _uploadBuffer.clear();
            break;
        }
        if (uploadResponse.statusCode != 202u)
        {
            return HresultFromGraphError(uploadResponse.statusCode, uploadResponse.body);
        }

        uint64_t nextOffset = 0u;
        hr                  = ParseNextExpectedUploadOffset(uploadResponse.body, _expectedSize, nextOffset);
        if (FAILED(hr) || nextOffset <= _uploadedBytes || nextOffset > endOffset + 1u)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (nextOffset < endOffset + 1u && ++partialAcknowledgements > 16u)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const size_t acceptedBytes = static_cast<size_t>(nextOffset - _uploadedBytes);
        _uploadedBytes             = nextOffset;
        _lastUploadStatus          = uploadResponse.statusCode;
        _uploadBuffer.erase(_uploadBuffer.begin(), _uploadBuffer.begin() + static_cast<std::ptrdiff_t>(acceptedBytes));
    }
    return S_OK;
}

HRESULT MicrosoftDriveFileWriter::WriteStreaming(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept
{
    if (! _hasExpectedSize || _position > _expectedSize || static_cast<uint64_t>(bytesToWrite) > (_expectedSize - _position))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    HRESULT hr = InitializeStreamingUpload();
    if (FAILED(hr))
    {
        return hr;
    }

    const auto* input = static_cast<const std::byte*>(buffer);
    while (*bytesWritten < bytesToWrite)
    {
        const size_t available = _uploadChunkBytes - _uploadBuffer.size();
        const size_t remaining = static_cast<size_t>(bytesToWrite - *bytesWritten);
        const size_t toCopy    = (std::min)(available, remaining);
        const size_t oldSize   = _uploadBuffer.size();
        _uploadBuffer.resize(oldSize + toCopy);
        std::memcpy(_uploadBuffer.data() + oldSize, input + *bytesWritten, toCopy);
        *bytesWritten += static_cast<unsigned long>(toCopy);
        _position += static_cast<uint64_t>(toCopy);

        if (_uploadBuffer.size() == _uploadChunkBytes || _position == _expectedSize)
        {
            hr = FlushStreamingChunk();
            if (FAILED(hr))
            {
                return hr;
            }
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
    hr = SendAuthenticatedGraphHttpRequest(_fileSystem->SnapshotSettings(),
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
    auto clearSessionResponseBody = wil::scope_exit([&] { SecureClear(sessionResponse.body); });
    hr = SendGraphJsonRequest(*_fileSystem, context, L"POST", BuildGraphCreateUploadSessionUrl(context, drivePath), bodyUtf8, {}, false, sessionResponse);
    if (FAILED(hr))
    {
        return hr;
    }
    if (sessionResponse.statusCode < 200u || sessionResponse.statusCode >= 300u)
    {
        return HresultFromGraphError(sessionResponse.statusCode, sessionResponse.body);
    }

    ValidatedPreauthenticatedUploadUrl uploadUrl{};
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

    std::vector<std::byte> buffer(static_cast<size_t>(chunkBytes));
    uint64_t offset                      = 0;
    unsigned int partialAcknowledgements = 0u;
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
        hr = SendPreauthenticatedUploadRequest(settings, L"PUT", uploadUrl, headers, buffer.data(), read, true, uploadResponse);
        if (FAILED(hr))
        {
            return hr;
        }

        if (uploadResponse.statusCode == 200u || uploadResponse.statusCode == 201u)
        {
            return endOffset + 1u == fileSize ? S_OK : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (uploadResponse.statusCode != 202u)
        {
            return HresultFromGraphError(uploadResponse.statusCode, uploadResponse.body);
        }

        uint64_t nextOffset = 0u;
        hr                  = ParseNextExpectedUploadOffset(uploadResponse.body, fileSize, nextOffset);
        if (FAILED(hr) || nextOffset <= offset || nextOffset > endOffset + 1u)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (nextOffset < endOffset + 1u)
        {
            if (++partialAcknowledgements > 16u)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }
        else
        {
            partialAcknowledgements = 0u;
        }

        offset = nextOffset;
        if (offset != endOffset + 1u)
        {
            LARGE_INTEGER position{};
            position.QuadPart = static_cast<LONGLONG>(offset);
            if (SetFilePointerEx(_tempFile.get(), position, nullptr, FILE_BEGIN) == FALSE)
            {
                return HRESULT_FROM_WIN32(GetLastError());
            }
        }
    }

    return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
}

HRESULT MicrosoftDriveFileWriter::Commit() noexcept
{
    if (_committed)
    {
        return S_OK;
    }
    if (_streaming)
    {
        if (! _hasExpectedSize || _position != _expectedSize)
        {
            return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        }
        HRESULT hr = FlushStreamingChunk();
        if (FAILED(hr))
        {
            return hr;
        }
        if (_uploadedBytes != _expectedSize || (_lastUploadStatus != 200u && _lastUploadStatus != 201u))
        {
            return HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY);
        }

        _committed = true;
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

#if defined(_DEBUG)
    if (g_debugMicrosoftDriveBypassAccessTokenForSelfTest.load())
    {
        accessTokenOut = "microsoft-drive-selftest-token";
        return S_OK;
    }
#endif

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

#if defined(_DEBUG)
namespace
{
struct DebugDriveItem
{
    std::wstring id;
    std::wstring name;
    std::wstring parentId;
    bool isFolder = false;
};

struct DebugDriveRequest
{
    std::wstring method;
    std::wstring url;
    bool allowRetry  = false;
    DWORD statusCode = 0;
};

[[nodiscard]] std::string JsonQuote(std::wstring_view value) noexcept
{
    std::string utf8 = Utf8FromUtf16(value);
    std::string quoted;
    quoted.reserve(utf8.size() + 2u);
    quoted.push_back('"');
    for (const char ch : utf8)
    {
        switch (ch)
        {
            case '\\': quoted.append("\\\\"); break;
            case '"': quoted.append("\\\""); break;
            case '\b': quoted.append("\\b"); break;
            case '\f': quoted.append("\\f"); break;
            case '\n': quoted.append("\\n"); break;
            case '\r': quoted.append("\\r"); break;
            case '\t': quoted.append("\\t"); break;
            default: quoted.push_back(ch); break;
        }
    }
    quoted.push_back('"');
    return quoted;
}

[[nodiscard]] std::wstring EnsureLeadingSlash(std::wstring path) noexcept
{
    if (path.empty() || path.front() != L'/')
    {
        path.insert(path.begin(), L'/');
    }
    return NormalizePluginPath(path);
}

[[nodiscard]] std::wstring ExtractRootPathFromDebugUrl(std::wstring_view url) noexcept
{
    const size_t root = url.find(L"/root");
    if (root == std::wstring_view::npos)
    {
        return {};
    }

    const size_t afterRoot = root + std::wstring_view(L"/root").size();
    if (afterRoot >= url.size() || url[afterRoot] == L'?' || url.substr(afterRoot).starts_with(L"/children"))
    {
        return L"/";
    }

    constexpr std::wstring_view kPathPrefix = L":/";
    if (! url.substr(afterRoot).starts_with(kPathPrefix))
    {
        return {};
    }

    const size_t pathStart = afterRoot + kPathPrefix.size();
    const size_t pathEnd   = url.find(L':', pathStart);
    if (pathEnd == std::wstring_view::npos || pathEnd < pathStart)
    {
        return {};
    }

    return EnsureLeadingSlash(std::wstring(url.substr(pathStart, pathEnd - pathStart)));
}

[[nodiscard]] std::wstring ExtractItemIdFromDebugUrl(std::wstring_view url) noexcept
{
    constexpr std::wstring_view kItems = L"/items/";
    const size_t itemStart             = url.find(kItems);
    if (itemStart == std::wstring_view::npos)
    {
        return {};
    }

    const size_t idStart = itemStart + kItems.size();
    size_t idEnd         = url.find(L'/', idStart);
    if (idEnd == std::wstring_view::npos)
    {
        idEnd = url.find(L'?', idStart);
    }
    if (idEnd == std::wstring_view::npos)
    {
        idEnd = url.size();
    }
    return std::wstring(url.substr(idStart, idEnd - idStart));
}

[[nodiscard]] bool TryReadDebugMoveBody(std::string_view bodyUtf8, std::wstring& nameOut, std::wstring& parentIdOut) noexcept
{
    nameOut.clear();
    parentIdOut.clear();

    yyjson_doc* doc = yyjson_read(bodyUtf8.data(), bodyUtf8.size(), YYJSON_READ_ALLOW_BOM);
    if (! doc)
    {
        return false;
    }
    Common::Json::UniqueDocument docOwner{doc};

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return false;
    }

    nameOut = TryGetJsonString(root, "name").value_or(L"");
    if (yyjson_val* parent = yyjson_obj_get(root, "parentReference"); parent && yyjson_is_obj(parent))
    {
        parentIdOut = TryGetJsonString(parent, "id").value_or(L"");
    }
    return ! nameOut.empty();
}

class DebugGraphDrive final
{
public:
    DebugGraphDrive()
    {
        _items.push_back(DebugDriveItem{.id = L"root", .name = L"", .parentId = L"", .isFolder = true});
    }

    DebugGraphDrive(const DebugGraphDrive&)            = delete;
    DebugGraphDrive& operator=(const DebugGraphDrive&) = delete;
    DebugGraphDrive(DebugGraphDrive&&)                 = delete;
    DebugGraphDrive& operator=(DebugGraphDrive&&)      = delete;

    DebugDriveItem* AddFolder(std::wstring_view path)
    {
        return AddItem(path, true);
    }

    DebugDriveItem* AddFile(std::wstring_view path)
    {
        return AddItem(path, false);
    }

    DebugDriveItem* AddRawChild(std::wstring_view parentPath, std::wstring_view name, bool isFolder)
    {
        DebugDriveItem* parent = FindByPath(parentPath);
        if (! parent || ! parent->isFolder)
        {
            return nullptr;
        }

        _items.push_back(DebugDriveItem{.id = std::format(L"id-{}", _nextId++), .name = std::wstring(name), .parentId = parent->id, .isFolder = isFolder});
        return &_items.back();
    }

    [[nodiscard]] bool Exists(std::wstring_view path, bool isFolder) noexcept
    {
        const DebugDriveItem* item = FindByPath(path);
        return item && item->isFolder == isFolder;
    }

    [[nodiscard]] std::wstring ItemId(std::wstring_view path) noexcept
    {
        const DebugDriveItem* item = FindByPath(path);
        return item ? item->id : std::wstring{};
    }

    [[nodiscard]] bool HasItemId(std::wstring_view id) const noexcept
    {
        return FindById(id) != nullptr;
    }

    [[nodiscard]] bool HasExactLeafName(std::wstring_view path, std::wstring_view expectedName) noexcept
    {
        const DebugDriveItem* item = FindByPath(path);
        return item != nullptr && item->name == expectedName;
    }

    [[nodiscard]] unsigned int CountRequests(std::wstring_view method) const noexcept
    {
        unsigned int count = 0;
        for (const DebugDriveRequest& request : requests)
        {
            if (OrdinalString::EqualsNoCase(request.method, method))
            {
                ++count;
            }
        }
        return count;
    }

    [[nodiscard]] bool AllRequestsAllowedRetry(std::wstring_view method) const noexcept
    {
        bool saw = false;
        for (const DebugDriveRequest& request : requests)
        {
            if (OrdinalString::EqualsNoCase(request.method, method))
            {
                saw = true;
                if (! request.allowRetry)
                {
                    return false;
                }
            }
        }
        return saw;
    }

    [[nodiscard]] bool AllRequestsDisallowedRetry(std::wstring_view method) const noexcept
    {
        bool saw = false;
        for (const DebugDriveRequest& request : requests)
        {
            if (OrdinalString::EqualsNoCase(request.method, method))
            {
                saw = true;
                if (request.allowRetry)
                {
                    return false;
                }
            }
        }
        return saw;
    }

    [[nodiscard]] HRESULT Handle(
        std::wstring_view method, std::wstring_view url, std::string_view bodyUtf8, bool allowRetry, HttpResponse& responseOut) noexcept
    {
        responseOut = {};
        if (OrdinalString::EqualsNoCase(method, L"GET"))
        {
            if (url.find(L"/children") != std::wstring_view::npos)
            {
                HandleChildren(url, responseOut);
            }
            else
            {
                HandleMetadata(url, responseOut);
            }
        }
        else if (OrdinalString::EqualsNoCase(method, L"PATCH"))
        {
            HandlePatch(url, bodyUtf8, responseOut);
        }
        else if (OrdinalString::EqualsNoCase(method, L"DELETE"))
        {
            HandleDelete(url, responseOut);
        }
        else if (OrdinalString::EqualsNoCase(method, L"POST"))
        {
            HandlePost(url, bodyUtf8, responseOut);
        }
        else
        {
            responseOut.statusCode = 400u;
            responseOut.body       = R"json({"error":{"code":"badRequest"}})json";
        }

        requests.push_back(
            DebugDriveRequest{.method = std::wstring(method), .url = std::wstring(url), .allowRetry = allowRetry, .statusCode = responseOut.statusCode});
        return S_OK;
    }

    int patchThrottleRemaining       = 0;
    int deleteThrottleRemaining      = 0;
    int postAmbiguousCreateRemaining = 0;
    std::vector<DebugDriveRequest> requests;

private:
    DebugDriveItem* AddItem(std::wstring_view rawPath, bool isFolder)
    {
        const std::wstring path = TrimTrailingSlashPreserveRoot(NormalizePluginPath(rawPath));
        if (path == L"/" || path.empty())
        {
            return FindById(L"root");
        }

        const size_t slash            = path.find_last_of(L'/');
        const std::wstring parentPath = slash == 0u ? L"/" : path.substr(0, slash);
        const std::wstring name       = path.substr(slash + 1u);
        DebugDriveItem* parent        = AddItem(parentPath, true);
        if (! parent)
        {
            return nullptr;
        }

        if (DebugDriveItem* existing = FindChild(parent->id, name))
        {
            existing->isFolder = isFolder;
            return existing;
        }

        _items.push_back(DebugDriveItem{.id = std::format(L"id-{}", _nextId++), .name = name, .parentId = parent->id, .isFolder = isFolder});
        return &_items.back();
    }

    [[nodiscard]] DebugDriveItem* FindById(std::wstring_view id) noexcept
    {
        for (DebugDriveItem& item : _items)
        {
            if (OrdinalString::EqualsNoCase(item.id, id))
            {
                return &item;
            }
        }
        return nullptr;
    }

    [[nodiscard]] const DebugDriveItem* FindById(std::wstring_view id) const noexcept
    {
        for (const DebugDriveItem& item : _items)
        {
            if (OrdinalString::EqualsNoCase(item.id, id))
            {
                return &item;
            }
        }
        return nullptr;
    }

    [[nodiscard]] DebugDriveItem* FindChild(std::wstring_view parentId, std::wstring_view name) noexcept
    {
        for (DebugDriveItem& item : _items)
        {
            if (OrdinalString::EqualsNoCase(item.parentId, parentId) && OrdinalString::EqualsNoCase(item.name, name))
            {
                return &item;
            }
        }
        return nullptr;
    }

    [[nodiscard]] const DebugDriveItem* FindChild(std::wstring_view parentId, std::wstring_view name) const noexcept
    {
        for (const DebugDriveItem& item : _items)
        {
            if (OrdinalString::EqualsNoCase(item.parentId, parentId) && OrdinalString::EqualsNoCase(item.name, name))
            {
                return &item;
            }
        }
        return nullptr;
    }

    [[nodiscard]] DebugDriveItem* FindByPath(std::wstring_view rawPath) noexcept
    {
        const std::wstring path = TrimTrailingSlashPreserveRoot(NormalizePluginPath(rawPath));
        if (path == L"/" || path.empty())
        {
            return FindById(L"root");
        }

        DebugDriveItem* current = FindById(L"root");
        size_t start            = 1u;
        while (current && start < path.size())
        {
            const size_t slash      = path.find(L'/', start);
            const size_t segmentEnd = slash == std::wstring::npos ? path.size() : slash;
            const std::wstring name = path.substr(start, segmentEnd - start);
            current                 = FindChild(current->id, name);
            start                   = slash == std::wstring::npos ? path.size() : slash + 1u;
        }
        return current;
    }

    [[nodiscard]] const DebugDriveItem* FindByPath(std::wstring_view rawPath) const noexcept
    {
        const std::wstring path = TrimTrailingSlashPreserveRoot(NormalizePluginPath(rawPath));
        if (path == L"/" || path.empty())
        {
            return FindById(L"root");
        }

        const DebugDriveItem* current = FindById(L"root");
        size_t start                  = 1u;
        while (current && start < path.size())
        {
            const size_t slash      = path.find(L'/', start);
            const size_t segmentEnd = slash == std::wstring::npos ? path.size() : slash;
            const std::wstring name = path.substr(start, segmentEnd - start);
            current                 = FindChild(current->id, name);
            start                   = slash == std::wstring::npos ? path.size() : slash + 1u;
        }
        return current;
    }

    [[nodiscard]] std::string ItemJson(const DebugDriveItem& item) const noexcept
    {
        return std::format(R"json({{"id":{},"name":{},"size":0,{}}})json",
                           JsonQuote(item.id),
                           JsonQuote(item.name),
                           item.isFolder ? R"json("folder":{})json" : R"json("file":{})json");
    }

    [[nodiscard]] std::string ChildrenJson(std::wstring_view parentId) const noexcept
    {
        std::string json = R"json({"value":[)json";
        bool first       = true;
        for (const DebugDriveItem& item : _items)
        {
            if (! OrdinalString::EqualsNoCase(item.parentId, parentId))
            {
                continue;
            }
            if (! first)
            {
                json.push_back(',');
            }
            first = false;
            json.append(ItemJson(item));
        }
        json.append("]}");
        return json;
    }

    void SetNotFound(HttpResponse& responseOut) const noexcept
    {
        responseOut.statusCode = 404u;
        responseOut.body       = R"json({"error":{"code":"itemNotFound"}})json";
    }

    void HandleMetadata(std::wstring_view url, HttpResponse& responseOut) const noexcept
    {
        const std::wstring path    = ExtractRootPathFromDebugUrl(url);
        const DebugDriveItem* item = FindByPath(path);
        if (! item)
        {
            SetNotFound(responseOut);
            return;
        }
        responseOut.statusCode = 200u;
        responseOut.body       = ItemJson(*item);
    }

    void HandleChildren(std::wstring_view url, HttpResponse& responseOut) const noexcept
    {
        const std::wstring path    = ExtractRootPathFromDebugUrl(url);
        const DebugDriveItem* item = FindByPath(path);
        if (! item)
        {
            SetNotFound(responseOut);
            return;
        }
        if (! item->isFolder)
        {
            responseOut.statusCode = 400u;
            responseOut.body       = R"json({"error":{"code":"notAFolder"}})json";
            return;
        }
        responseOut.statusCode = 200u;
        responseOut.body       = ChildrenJson(item->id);
    }

    void HandlePatch(std::wstring_view url, std::string_view bodyUtf8, HttpResponse& responseOut) noexcept
    {
        if (patchThrottleRemaining > 0)
        {
            --patchThrottleRemaining;
            responseOut.statusCode = 429u;
            responseOut.retryAfter = L"0";
            responseOut.body       = R"json({"error":{"code":"tooManyRequests"}})json";
            return;
        }

        DebugDriveItem* item = FindById(ExtractItemIdFromDebugUrl(url));
        if (! item)
        {
            SetNotFound(responseOut);
            return;
        }

        std::wstring name;
        std::wstring parentId;
        if (! TryReadDebugMoveBody(bodyUtf8, name, parentId))
        {
            responseOut.statusCode = 400u;
            responseOut.body       = R"json({"error":{"code":"badRequest"}})json";
            return;
        }

        if (! parentId.empty())
        {
            DebugDriveItem* parent = FindById(parentId);
            if (! parent || ! parent->isFolder)
            {
                SetNotFound(responseOut);
                return;
            }
            item->parentId = parentId;
        }
        item->name = name;

        responseOut.statusCode = 200u;
        responseOut.body       = ItemJson(*item);
    }

    void HandleDelete(std::wstring_view url, HttpResponse& responseOut) noexcept
    {
        if (deleteThrottleRemaining > 0)
        {
            --deleteThrottleRemaining;
            responseOut.statusCode = 429u;
            responseOut.retryAfter = L"0";
            responseOut.body       = R"json({"error":{"code":"tooManyRequests"}})json";
            return;
        }

        const std::wstring id = ExtractItemIdFromDebugUrl(url);
        const auto it =
            std::find_if(_items.begin(), _items.end(), [&](const DebugDriveItem& item) noexcept { return OrdinalString::EqualsNoCase(item.id, id); });
        if (it == _items.end())
        {
            SetNotFound(responseOut);
            return;
        }

        _items.erase(it);
        responseOut.statusCode = 204u;
    }

    void HandlePost(std::wstring_view url, std::string_view bodyUtf8, HttpResponse& responseOut) noexcept
    {
        const std::wstring parentId = ExtractItemIdFromDebugUrl(url);
        DebugDriveItem* parent      = FindById(parentId);
        if (! parent || ! parent->isFolder)
        {
            SetNotFound(responseOut);
            return;
        }

        std::wstring name;
        std::wstring ignoredParentId;
        if (! TryReadDebugMoveBody(bodyUtf8, name, ignoredParentId))
        {
            responseOut.statusCode = 400u;
            responseOut.body       = R"json({"error":{"code":"badRequest"}})json";
            return;
        }

        if (FindChild(parent->id, name))
        {
            responseOut.statusCode = 409u;
            responseOut.body       = R"json({"error":{"code":"nameAlreadyExists"}})json";
            return;
        }

        _items.push_back(DebugDriveItem{.id = std::format(L"id-{}", _nextId++), .name = name, .parentId = parent->id, .isFolder = true});
        DebugDriveItem& created = _items.back();

        if (postAmbiguousCreateRemaining > 0)
        {
            --postAmbiguousCreateRemaining;
            responseOut.statusCode = 503u;
            responseOut.retryAfter = L"0";
            responseOut.body       = R"json({"error":{"code":"serviceUnavailable"}})json";
            return;
        }

        responseOut.statusCode = 201u;
        responseOut.body       = ItemJson(created);
    }

    std::vector<DebugDriveItem> _items;
    unsigned int _nextId = 1;
};

[[nodiscard]] HRESULT DebugGraphHook(void* cookie,
                                     std::wstring_view method,
                                     std::wstring_view url,
                                     [[maybe_unused]] std::span<const HttpHeader> headers,
                                     std::string_view bodyUtf8,
                                     bool allowRetry,
                                     HttpResponse& responseOut) noexcept
{
    if (! cookie)
    {
        return E_POINTER;
    }
    return static_cast<DebugGraphDrive*>(cookie)->Handle(method, url, bodyUtf8, allowRetry, responseOut);
}

[[nodiscard]] DriveContext MakeDebugDriveContext() noexcept
{
    DriveContext context{};
    context.connectionName      = L"microsoft-drive-selftest";
    context.profile.name        = context.connectionName;
    context.profile.pluginId    = L"builtin/file-system-onedrive-personal";
    context.profile.authMode    = L"oauth2Pkce";
    context.authority           = L"consumers";
    context.scopeText           = L"offline_access Files.ReadWrite User.Read openid profile";
    context.driveId             = L"drive-selftest";
    context.driveDisplayName    = L"Microsoft Drive SelfTest";
    context.driveVolumeLabel    = L"Microsoft Drive SelfTest";
    context.persistRefreshToken = false;
    return context;
}

constexpr Common::DebugSelfTest::Check DebugCheck{L"Microsoft Drive"};

[[nodiscard]] wil::com_ptr<FileSystemMicrosoftDrive> MakeDebugFileSystem() noexcept
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs;
    auto* raw = new (std::nothrow) FileSystemMicrosoftDrive(FileSystemMicrosoftDriveMode::OneDrivePersonal, nullptr);
    if (raw)
    {
        fs.attach(raw);
    }
    return fs;
}

void RunDebugPatchRetryMergeSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate Microsoft Drive instance", passed, failed))
    {
        return;
    }
    const DriveContext context = MakeDebugDriveContext();
    DebugGraphDrive graph;
    graph.AddFolder(L"/src/Foo");
    graph.AddFile(L"/src/Foo/child.txt");
    graph.AddFolder(L"/dst/Foo");
    graph.patchThrottleRemaining = 1;

    DebugHttpRequestHookScope hook(DebugGraphHook, &graph);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope noRetrySleep(g_debugMicrosoftDriveSuppressRetrySleepForSelfTest, true);

    const HRESULT hr = MoveOrRenameItem(*fs, context, context, L"/src/Foo", L"/dst/Foo", FILESYSTEM_FLAG_RECURSIVE);
    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), L"throttled child PATCH merge should retain the source folder as cleanup debt", passed, failed);
    DebugCheck(graph.CountRequests(L"PATCH") == 2u, L"throttled child PATCH should be attempted twice", passed, failed);
    DebugCheck(graph.AllRequestsAllowedRetry(L"PATCH"), L"PATCH Graph mutations should enable retry handling", passed, failed);
    DebugCheck(graph.Exists(L"/dst/Foo/child.txt", false), L"retried child move should land in destination", passed, failed);
    DebugCheck(! graph.Exists(L"/src/Foo/child.txt", false), L"retried child move should leave no source child behind", passed, failed);
    DebugCheck(graph.Exists(L"/src/Foo", true), L"merged source folder should remain because Graph cannot delete it atomically only-if-empty", passed, failed);
    DebugCheck(graph.CountRequests(L"DELETE") == 0u, L"directory merge should never issue recursive source-folder DELETE", passed, failed);
}

void RunDebugCaseOnlyRenameSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"case-only rename selftest should allocate Microsoft Drive instance", passed, failed))
    {
        return;
    }

    const DriveContext context = MakeDebugDriveContext();
    DebugGraphDrive graph;
    graph.AddFile(L"/CaseOnly.txt");

    DebugHttpRequestHookScope hook(DebugGraphHook, &graph);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope noRetrySleep(g_debugMicrosoftDriveSuppressRetrySleepForSelfTest, true);

    const HRESULT hr = MoveOrRenameItem(*fs, context, context, L"/CaseOnly.txt", L"/caseonly.txt", FILESYSTEM_FLAG_NONE);
    DebugCheck(hr == S_OK, L"Microsoft Drive case-only rename should succeed", passed, failed);
    DebugCheck(graph.CountRequests(L"PATCH") == 1u, L"Microsoft Drive case-only rename should issue one Graph PATCH", passed, failed);
    DebugCheck(
        graph.HasExactLeafName(L"/caseonly.txt", L"caseonly.txt"), L"Microsoft Drive case-only rename should update the stored leaf casing", passed, failed);
}

void RunDebugDeleteRetrySelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate Microsoft Drive instance", passed, failed))
    {
        return;
    }
    const DriveContext context = MakeDebugDriveContext();
    DebugGraphDrive graph;
    graph.AddFile(L"/delete-me.txt");
    graph.deleteThrottleRemaining = 1;

    const std::wstring itemId = graph.ItemId(L"/delete-me.txt");
    DebugHttpRequestHookScope hook(DebugGraphHook, &graph);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope noRetrySleep(g_debugMicrosoftDriveSuppressRetrySleepForSelfTest, true);

    const HRESULT hr = DeleteItemById(*fs, context, itemId);
    DebugCheck(hr == S_OK, L"throttled DELETE should complete", passed, failed);
    DebugCheck(graph.CountRequests(L"DELETE") == 2u, L"throttled DELETE should be attempted twice", passed, failed);
    DebugCheck(graph.AllRequestsAllowedRetry(L"DELETE"), L"DELETE Graph mutations should enable retry handling", passed, failed);
    DebugCheck(! graph.Exists(L"/delete-me.txt", false), L"retried delete should remove the item", passed, failed);
}

void RunDebugPostCreateReconcileSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate Microsoft Drive instance", passed, failed))
    {
        return;
    }
    const DriveContext context = MakeDebugDriveContext();
    DebugGraphDrive graph;
    graph.AddFolder(L"/parent");
    graph.postAmbiguousCreateRemaining = 1;

    DebugHttpRequestHookScope hook(DebugGraphHook, &graph);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope noRetrySleep(g_debugMicrosoftDriveSuppressRetrySleepForSelfTest, true);

    const HRESULT hr = CreateDirectoryItem(*fs, context, L"/parent/new-folder");
    DebugCheck(hr == S_OK, L"ambiguous POST create should reconcile committed directory", passed, failed);
    DebugCheck(graph.CountRequests(L"POST") == 1u, L"ambiguous committed POST create should not be blindly replayed", passed, failed);
    DebugCheck(graph.AllRequestsDisallowedRetry(L"POST"), L"POST create should use explicit reconciliation instead of blind transport retry", passed, failed);
    DebugCheck(graph.Exists(L"/parent/new-folder", true), L"ambiguous POST create should leave created folder visible", passed, failed);
}

void RunDebugTypeMismatchPartialMergeSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate Microsoft Drive instance", passed, failed))
    {
        return;
    }
    const DriveContext context = MakeDebugDriveContext();
    DebugGraphDrive graph;
    graph.AddFolder(L"/src/Foo");
    graph.AddFolder(L"/src/Foo/conflict");
    graph.AddFile(L"/src/Foo/ok.txt");
    graph.AddFolder(L"/dst/Foo");
    graph.AddFile(L"/dst/Foo/conflict");

    unsigned int prompts             = 0;
    const MoveIssueReporter reporter = [&](const wchar_t*, const wchar_t*, HRESULT status, FileSystemIssueAction& action) noexcept -> HRESULT
    {
        ++prompts;
        if (status != HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS))
        {
            return E_FAIL;
        }
        action = FileSystemIssueAction::Overwrite;
        return S_OK;
    };

    DebugHttpRequestHookScope hook(DebugGraphHook, &graph);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope noRetrySleep(g_debugMicrosoftDriveSuppressRetrySleepForSelfTest, true);

    const HRESULT hr = MoveOrRenameItem(*fs, context, context, L"/src/Foo", L"/dst/Foo", FILESYSTEM_FLAG_RECURSIVE, reporter);
    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), L"type-mismatch overwrite should finish as partial copy", passed, failed);
    DebugCheck(prompts == 1u, L"type-mismatch merge should prompt exactly once", passed, failed);
    DebugCheck(graph.Exists(L"/src/Foo", true), L"partial merge should preserve source folder", passed, failed);
    DebugCheck(graph.Exists(L"/src/Foo/conflict", true), L"unresolvable type-mismatch source child should remain", passed, failed);
    DebugCheck(! graph.Exists(L"/src/Foo/ok.txt", false), L"non-conflicting sibling should still move", passed, failed);
    DebugCheck(graph.Exists(L"/dst/Foo/ok.txt", false), L"non-conflicting sibling should land in destination", passed, failed);
    DebugCheck(graph.Exists(L"/dst/Foo/conflict", false), L"type-mismatch destination child should remain unchanged", passed, failed);
}

void RunDebugDescendantMoveGuardSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate Microsoft Drive instance for descendant guard", passed, failed))
    {
        return;
    }
    const DriveContext context = MakeDebugDriveContext();
    DebugGraphDrive graph;
    graph.AddFolder(L"/src/Foo");
    graph.AddFolder(L"/src/Foo/Sub");

    DebugHttpRequestHookScope hook(DebugGraphHook, &graph);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope noRetrySleep(g_debugMicrosoftDriveSuppressRetrySleepForSelfTest, true);

    const HRESULT hr = MoveOrRenameItem(*fs, context, context, L"/src/Foo", L"/src/Foo/Sub", FILESYSTEM_FLAG_RECURSIVE);
    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_INVALID_PARAMETER), L"moving a folder into its own descendant should be rejected", passed, failed);
    DebugCheck(graph.Exists(L"/src/Foo", true), L"descendant-guard failure should preserve source folder", passed, failed);
    DebugCheck(graph.Exists(L"/src/Foo/Sub", true), L"descendant-guard failure should preserve descendant folder", passed, failed);
}

void RunDebugParseChildrenSkipsEmptyNameSelfTest(unsigned int& passed, unsigned int& failed)
{
    std::vector<FilesInformationMicrosoftDrive::Entry> entries;
    std::wstring nextLink;
    bool incompleteDueToInvalidChildName = false;
    constexpr std::string_view body      = R"json({"value":[{"id":"bad","name":"","folder":{}},{"id":"ok","name":"ok.txt","size":3,"file":{}}]})json";

    const HRESULT hr = ParseChildren(body, entries, nextLink, &incompleteDueToInvalidChildName);
    DebugCheck(hr == S_OK, L"ParseChildren should accept a page containing an empty-name child", passed, failed);
    DebugCheck(incompleteDueToInvalidChildName, L"ParseChildren should mark pages containing empty-name children as incomplete", passed, failed);
    DebugCheck(entries.size() == 1u, L"ParseChildren should skip empty-name children", passed, failed);
    if (entries.size() == 1u)
    {
        DebugCheck(entries[0].name == L"ok.txt", L"ParseChildren should preserve valid siblings after an empty-name child", passed, failed);
    }
}

void RunDebugEmptyNameChildBlocksSourceDeleteSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(
            static_cast<bool>(fs), L"debug selftest should allocate Microsoft Drive instance for empty-name child source-delete guard", passed, failed))
    {
        return;
    }
    const DriveContext context = MakeDebugDriveContext();
    DebugGraphDrive graph;
    graph.AddFolder(L"/src/Foo");
    graph.AddFile(L"/src/Foo/ok.txt");
    graph.AddFolder(L"/dst/Foo");
    DebugCheck(graph.AddRawChild(L"/src/Foo", L"", false) != nullptr, L"debug graph should add malformed empty-name source child", passed, failed);

    DebugHttpRequestHookScope hook(DebugGraphHook, &graph);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope noRetrySleep(g_debugMicrosoftDriveSuppressRetrySleepForSelfTest, true);

    const HRESULT hr = MoveOrRenameItem(*fs, context, context, L"/src/Foo", L"/dst/Foo", FILESYSTEM_FLAG_RECURSIVE);
    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), L"empty-name child merge should finish as partial copy", passed, failed);
    DebugCheck(graph.Exists(L"/src/Foo", true), L"empty-name child merge should preserve source folder", passed, failed);
    DebugCheck(! graph.Exists(L"/src/Foo/ok.txt", false), L"empty-name child merge should still move valid siblings", passed, failed);
    DebugCheck(graph.Exists(L"/dst/Foo/ok.txt", false), L"empty-name child merge should land valid siblings in destination", passed, failed);
}

struct MergeConcurrentAddContext final
{
    DebugGraphDrive* graph = nullptr;
    bool injected          = false;
};

// Inject a new source child after the originally-enumerated child commits at the destination. The merge
// must leave the source folder in place because Graph DELETE cannot atomically require it to remain empty.
[[nodiscard]] HRESULT MergeConcurrentAddHook(void* cookie,
                                             std::wstring_view method,
                                             std::wstring_view url,
                                             std::span<const HttpHeader> headers,
                                             std::string_view bodyUtf8,
                                             bool allowRetry,
                                             HttpResponse& responseOut) noexcept
{
    auto* ctx = static_cast<MergeConcurrentAddContext*>(cookie);
    if (! ctx || ! ctx->graph)
    {
        return E_POINTER;
    }

    const HRESULT hr = DebugGraphHook(ctx->graph, method, url, headers, bodyUtf8, allowRetry, responseOut);
    if (SUCCEEDED(hr) && OrdinalString::EqualsNoCase(method, L"PATCH") && ! ctx->injected)
    {
        ctx->graph->AddFile(L"/src/Foo/late.txt");
        ctx->injected = true;
    }
    return hr;
}

void RunDebugMergeNeverRecursivelyDeletesSourceFolderSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate Microsoft Drive instance for re-list drain guard", passed, failed))
    {
        return;
    }
    const DriveContext context = MakeDebugDriveContext();
    DebugGraphDrive graph;
    graph.AddFolder(L"/src/Foo");
    graph.AddFile(L"/src/Foo/ok.txt");
    graph.AddFolder(L"/dst/Foo");

    MergeConcurrentAddContext injection{};
    injection.graph = &graph;

    DebugHttpRequestHookScope hook(MergeConcurrentAddHook, &injection);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope noRetrySleep(g_debugMicrosoftDriveSuppressRetrySleepForSelfTest, true);

    const HRESULT hr = MoveOrRenameItem(*fs, context, context, L"/src/Foo", L"/dst/Foo", FILESYSTEM_FLAG_RECURSIVE);

    DebugCheck(injection.injected, L"merge race test should inject a child after the original child move", passed, failed);
    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY), L"safe merge should report retained source-folder cleanup", passed, failed);
    DebugCheck(graph.Exists(L"/src/Foo", true), L"safe merge should retain the source folder", passed, failed);
    DebugCheck(graph.Exists(L"/src/Foo/late.txt", false), L"safe merge should preserve the concurrently-added child", passed, failed);
    DebugCheck(graph.Exists(L"/dst/Foo/ok.txt", false), L"safe merge should still move the originally enumerated child", passed, failed);
    DebugCheck(! graph.Exists(L"/src/Foo/ok.txt", false), L"safe merge should remove the moved child from the source", passed, failed);
    DebugCheck(graph.CountRequests(L"DELETE") == 0u, L"safe merge should issue no recursive source-folder DELETE", passed, failed);
}

void RunDebugMoveRejectsInvalidOptionsSizeSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate Microsoft Drive instance for invalid Move options guard", passed, failed))
    {
        return;
    }

    FileSystemOptions badOptions{};
    badOptions.sizeBytes = sizeof(FileSystemOptions) - 1u;
    const HRESULT hr     = MoveSingleItemWithConflicts(*fs, L"/src.txt", L"/dst.txt", FILESYSTEM_FLAG_NONE, &badOptions, nullptr, nullptr);
    DebugCheck(hr == E_INVALIDARG, L"Microsoft Drive Move should reject invalid FileSystemOptions::sizeBytes before reading options", passed, failed);
}

void RunDebugMergeRecursionDepthCapSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate Microsoft Drive instance for recursion cap", passed, failed))
    {
        return;
    }
    const DriveContext context = MakeDebugDriveContext();
    DebugGraphDrive graph;

    std::wstring sourcePath = L"/src/Foo";
    std::wstring destPath   = L"/dst/Foo";
    graph.AddFolder(sourcePath);
    graph.AddFolder(destPath);
    for (unsigned int depth = 0; depth < 70u; ++depth)
    {
        const std::wstring leaf = std::format(L"d{}", depth);
        sourcePath              = JoinPath(sourcePath, leaf);
        destPath                = JoinPath(destPath, leaf);
        graph.AddFolder(sourcePath);
        graph.AddFolder(destPath);
    }
    graph.AddFile(JoinPath(sourcePath, L"leaf.txt"));

    DebugHttpRequestHookScope hook(DebugGraphHook, &graph);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope noRetrySleep(g_debugMicrosoftDriveSuppressRetrySleepForSelfTest, true);

    const HRESULT hr = MoveOrRenameItem(*fs, context, context, L"/src/Foo", L"/dst/Foo", FILESYSTEM_FLAG_RECURSIVE);
    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_STACK_OVERFLOW), L"deep Microsoft Drive merge recursion should trip the provider depth cap", passed, failed);
    DebugCheck(graph.Exists(JoinPath(sourcePath, L"leaf.txt"), false), L"recursion-cap failure should preserve the deep source leaf", passed, failed);
}

void RunDebugOverwriteCancelRestoresDestinationSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"debug selftest should allocate Microsoft Drive instance for overwrite rollback", passed, failed))
    {
        return;
    }

    const DriveContext context = MakeDebugDriveContext();
    DebugGraphDrive graph;
    graph.AddFile(L"/src.txt");
    graph.AddFile(L"/dst.txt");

    DebugHttpRequestHookScope hook(DebugGraphHook, &graph);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope noRetrySleep(g_debugMicrosoftDriveSuppressRetrySleepForSelfTest, true);

    const HRESULT hr =
        MoveOrRenameItem(*fs, context, context, L"/src.txt", L"/dst.txt", FILESYSTEM_FLAG_ALLOW_OVERWRITE, MoveIssueReporter{}, true, []() noexcept {
        return HRESULT_FROM_WIN32(ERROR_CANCELLED);
    });

    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_CANCELLED), L"overwrite move cancellation should propagate ERROR_CANCELLED", passed, failed);
    DebugCheck(graph.Exists(L"/src.txt", false), L"cancelled overwrite move should preserve the source", passed, failed);
    DebugCheck(graph.Exists(L"/dst.txt", false), L"cancelled overwrite move should restore the original destination", passed, failed);
}

struct CleanupFailureDebugContext final
{
    DebugGraphDrive* graph       = nullptr;
    unsigned int cleanupAttempts = 0u;
};

[[nodiscard]] HRESULT CleanupFailureDebugHook(void* cookie,
                                              std::wstring_view method,
                                              std::wstring_view url,
                                              std::span<const HttpHeader> headers,
                                              std::string_view bodyUtf8,
                                              bool allowRetry,
                                              HttpResponse& responseOut) noexcept
{
    auto* context = static_cast<CleanupFailureDebugContext*>(cookie);
    if (! context || ! context->graph)
    {
        return E_POINTER;
    }
    if (OrdinalString::EqualsNoCase(method, L"DELETE"))
    {
        ++context->cleanupAttempts;
        responseOut.statusCode = 500u;
        responseOut.body       = R"json({"error":{"code":"serverError"}})json";
        return S_OK;
    }
    return DebugGraphHook(context->graph, method, url, headers, bodyUtf8, allowRetry, responseOut);
}

void RunDebugCommittedMoveCleanupFailureIsWarningSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"cleanup-warning test should allocate Microsoft Drive instance", passed, failed))
    {
        return;
    }

    const DriveContext context = MakeDebugDriveContext();
    DebugGraphDrive graph;
    graph.AddFile(L"/src.txt");
    graph.AddFile(L"/dst.txt");
    const std::wstring sourceId         = graph.ItemId(L"/src.txt");
    const std::wstring oldDestinationId = graph.ItemId(L"/dst.txt");

    CleanupFailureDebugContext failure{.graph = &graph};
    DebugHttpRequestHookScope hook(CleanupFailureDebugHook, &failure);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope noRetrySleep(g_debugMicrosoftDriveSuppressRetrySleepForSelfTest, true);

    MoveCommitResult commit{};
    const HRESULT hr =
        MoveOrRenameItem(*fs, context, context, L"/src.txt", L"/dst.txt", FILESYSTEM_FLAG_ALLOW_OVERWRITE, MoveIssueReporter{}, true, CancelProbe{}, &commit);
    DebugCheck(hr == S_OK, L"committed move should remain successful when only backup cleanup fails", passed, failed);
    DebugCheck(commit.primaryMutationCommitted && FAILED(commit.cleanupStatus) && commit.rollbackStatus == S_OK,
               L"committed move should return structured cleanup debt without rollback failure",
               passed,
               failed);
    DebugCheck(! graph.Exists(L"/src.txt", false) && graph.ItemId(L"/dst.txt") == sourceId,
               L"cleanup-warning move should leave the requested source-to-destination mutation committed",
               passed,
               failed);
    DebugCheck(graph.HasItemId(oldDestinationId), L"failed cleanup should retain the recoverable overwrite backup", passed, failed);
    DebugCheck(failure.cleanupAttempts > 0u, L"cleanup-warning test should exercise backup deletion failure", passed, failed);
}

[[nodiscard]] std::wstring_view FindDebugHeader(std::span<const HttpHeader> headers, std::wstring_view name) noexcept
{
    const auto found =
        std::find_if(headers.begin(), headers.end(), [&](const HttpHeader& header) noexcept { return OrdinalString::EqualsNoCase(header.name, name); });
    return found == headers.end() ? std::wstring_view{} : std::wstring_view(found->value);
}

void RunDebugCredentialUrlValidationSelfTest(unsigned int& passed, unsigned int& failed)
{
    ValidatedGraphApiUrl graphUrl{};
    DebugCheck(SUCCEEDED(ValidateGraphApiUrl(L"https://graph.microsoft.com/v1.0/me/drive?$skiptoken=opaque", graphUrl)),
               L"Graph URL validation should accept the configured HTTPS v1.0 origin",
               passed,
               failed);
    DebugCheck(FAILED(ValidateGraphApiUrl(L"http://graph.microsoft.com/v1.0/me/drive", graphUrl)),
               L"Graph URL validation should reject plaintext HTTP",
               passed,
               failed);
    DebugCheck(FAILED(ValidateGraphApiUrl(L"https://graph.microsoft.com.evil.example/v1.0/me/drive", graphUrl)),
               L"Graph URL validation should reject foreign origins",
               passed,
               failed);
    DebugCheck(FAILED(ValidateGraphApiUrl(L"https://user@graph.microsoft.com/v1.0/me/drive", graphUrl)),
               L"Graph URL validation should reject userinfo",
               passed,
               failed);
    DebugCheck(FAILED(ValidateGraphApiUrl(L"https://graph.microsoft.com/v1.0/me/drive#fragment", graphUrl)),
               L"Graph URL validation should reject fragments",
               passed,
               failed);
    DebugCheck(FAILED(ValidateGraphApiUrl(L"https://graph.microsoft.com:444/v1.0/me/drive", graphUrl)),
               L"Graph URL validation should reject non-default ports",
               passed,
               failed);
    DebugCheck(FAILED(ValidateGraphApiUrl(L"https://graph.microsoft.com:invalid/v1.0/me/drive", graphUrl)),
               L"Graph URL validation should reject malformed ports",
               passed,
               failed);

    ValidatedPreauthenticatedUploadUrl uploadUrl{};
    DebugCheck(SUCCEEDED(ValidatePreauthenticatedUploadUrl(L"https://sn3302.up.1drv.com/up/session?opaque=1", uploadUrl)),
               L"upload URL validation should accept the documented OneDrive preauthenticated origin",
               passed,
               failed);
    DebugCheck(SUCCEEDED(ValidatePreauthenticatedUploadUrl(L"https://tenant.sharepoint.com/_api/upload?opaque=1", uploadUrl)),
               L"upload URL validation should accept a tenant SharePoint preauthenticated origin",
               passed,
               failed);
    DebugCheck(FAILED(ValidatePreauthenticatedUploadUrl(L"http://sn3302.up.1drv.com/up/session", uploadUrl)),
               L"upload URL validation should reject plaintext HTTP",
               passed,
               failed);
    DebugCheck(FAILED(ValidatePreauthenticatedUploadUrl(L"https://up.1drv.com.evil.example/up/session", uploadUrl)),
               L"upload URL validation should reject suffix-confusion origins",
               passed,
               failed);
    DebugCheck(FAILED(ValidatePreauthenticatedUploadUrl(L"https://user@sn3302.up.1drv.com/up/session", uploadUrl)),
               L"upload URL validation should reject userinfo",
               passed,
               failed);
    DebugCheck(FAILED(ValidatePreauthenticatedUploadUrl(L"https://sn3302.up.1drv.com/up/session#fragment", uploadUrl)),
               L"upload URL validation should reject fragments",
               passed,
               failed);
}

enum class ContinuationDebugScenario
{
    ApprovedSameOrigin,
    ForeignOrigin,
    PlaintextOrigin,
    Repeated,
};

struct ContinuationDebugContext final
{
    ContinuationDebugScenario scenario = ContinuationDebugScenario::ApprovedSameOrigin;
    unsigned int requestCount          = 0u;
};

[[nodiscard]] HRESULT ContinuationDebugHook(void* cookie,
                                            std::wstring_view method,
                                            [[maybe_unused]] std::wstring_view url,
                                            [[maybe_unused]] std::span<const HttpHeader> headers,
                                            [[maybe_unused]] std::string_view bodyUtf8,
                                            [[maybe_unused]] bool allowRetry,
                                            HttpResponse& responseOut) noexcept
{
    auto* context = static_cast<ContinuationDebugContext*>(cookie);
    if (context == nullptr || ! OrdinalString::EqualsNoCase(method, L"GET"))
    {
        return E_UNEXPECTED;
    }

    ++context->requestCount;
    responseOut.statusCode = 200u;
    if (context->requestCount == 1u)
    {
        std::string_view nextLink;
        switch (context->scenario)
        {
            case ContinuationDebugScenario::ApprovedSameOrigin:
            case ContinuationDebugScenario::Repeated: nextLink = "https://graph.microsoft.com/v1.0/drives/debug/root/children?$skiptoken=page-two"; break;
            case ContinuationDebugScenario::ForeignOrigin: nextLink = "https://foreign.example/v1.0/drives/debug/root/children?$skiptoken=stolen"; break;
            case ContinuationDebugScenario::PlaintextOrigin: nextLink = "http://graph.microsoft.com/v1.0/drives/debug/root/children?$skiptoken=stolen"; break;
        }
        responseOut.body = std::format(R"json({{"value":[],"@odata.nextLink":"{}"}})json", nextLink);
        return S_OK;
    }

    responseOut.body = context->scenario == ContinuationDebugScenario::Repeated
                           ? R"json({"value":[],"@odata.nextLink":"https://graph.microsoft.com/v1.0/drives/debug/root/children?$skiptoken=page-two"})json"
                           : R"json({"value":[]})json";
    return S_OK;
}

void RunDebugContinuationBoundarySelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"continuation boundary test should allocate Microsoft Drive instance", passed, failed))
    {
        return;
    }
    const DriveContext driveContext = MakeDebugDriveContext();
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);

    const auto runScenario = [&](ContinuationDebugScenario scenario, HRESULT expectedHr, unsigned int expectedRequests, std::wstring_view message)
    {
        ContinuationDebugContext context{.scenario = scenario};
        DebugHttpRequestHookScope hook(ContinuationDebugHook, &context);
        std::vector<FilesInformationMicrosoftDrive::Entry> entries;
        const HRESULT hr = ListDirectory(*fs, driveContext, L"/", entries);
        DebugCheck(hr == expectedHr && context.requestCount == expectedRequests, message.data(), passed, failed);
    };

    runScenario(ContinuationDebugScenario::ApprovedSameOrigin, S_OK, 2u, L"approved same-origin Graph pagination should issue the second request");
    runScenario(ContinuationDebugScenario::ForeignOrigin,
                HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                1u,
                L"foreign Graph continuation should fail before a second request");
    runScenario(ContinuationDebugScenario::PlaintextOrigin,
                HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED),
                1u,
                L"plaintext Graph continuation should fail before a second request");
    runScenario(
        ContinuationDebugScenario::Repeated, HRESULT_FROM_WIN32(ERROR_INVALID_DATA), 2u, L"repeated Graph continuation should fail before a third request");
}

struct CredentialDiagnosticDebugContext final
{
    unsigned int requestCount = 0u;
    std::wstring captured;
};

[[nodiscard]] HRESULT CredentialDiagnosticFailureHook(void* cookie,
                                                      [[maybe_unused]] std::wstring_view method,
                                                      [[maybe_unused]] std::wstring_view url,
                                                      [[maybe_unused]] std::span<const HttpHeader> headers,
                                                      [[maybe_unused]] std::string_view bodyUtf8,
                                                      [[maybe_unused]] bool allowRetry,
                                                      [[maybe_unused]] HttpResponse& responseOut) noexcept
{
    auto* context = static_cast<CredentialDiagnosticDebugContext*>(cookie);
    if (context == nullptr)
    {
        return E_POINTER;
    }
    ++context->requestCount;
    return HRESULT_FROM_WIN32(ERROR_WINHTTP_CANNOT_CONNECT);
}

void CredentialDiagnosticCapture(void* cookie, std::wstring_view diagnostic) noexcept
{
    auto* context = static_cast<CredentialDiagnosticDebugContext*>(cookie);
    if (context)
    {
        context->captured.append(diagnostic);
        context->captured.push_back(L'\n');
    }
}

void RunDebugCredentialDiagnosticRedactionSelfTest(unsigned int& passed, unsigned int& failed)
{
    constexpr char kBearerSentinel[]            = "OBSERVATORY_BEARER_SENTINEL";
    constexpr std::wstring_view kGraphSentinel  = L"OBSERVATORY_GRAPH_QUERY_SENTINEL";
    constexpr std::wstring_view kUploadSentinel = L"OBSERVATORY_UPLOAD_QUERY_SENTINEL";
    CredentialDiagnosticDebugContext context{};
    DebugHttpRequestHookScope requestHook(CredentialDiagnosticFailureHook, &context);
    DebugHttpDiagnosticHookScope diagnosticHook(CredentialDiagnosticCapture, &context);

    FileSystemMicrosoftDrive::Settings settings{};
    HttpResponse response{};
    const std::wstring graphUrl = std::format(L"https://graph.microsoft.com/v1.0/me/drive?$skiptoken={}", kGraphSentinel);
    HRESULT hr                  = SendAuthenticatedGraphHttpRequest(settings, L"GET", graphUrl, kBearerSentinel, {}, nullptr, 0u, nullptr, false, response);
    DebugCheck(FAILED(hr), L"injected authorized Graph transport failure should propagate", passed, failed);

    ValidatedPreauthenticatedUploadUrl uploadUrl{};
    const std::wstring rawUploadUrl = std::format(L"https://sn3302.up.1drv.com/up/session?auth={}", kUploadSentinel);
    hr                              = ValidatePreauthenticatedUploadUrl(rawUploadUrl, uploadUrl);
    if (SUCCEEDED(hr))
    {
        hr = SendPreauthenticatedUploadRequest(settings, L"PUT", uploadUrl, {}, nullptr, 0u, false, response);
    }
    DebugCheck(FAILED(hr), L"injected preauthenticated upload transport failure should propagate", passed, failed);
    DebugCheck(context.requestCount == 2u, L"diagnostic redaction test should inject both transport failures", passed, failed);
    DebugCheck(context.captured.find(Utf16FromUtf8(kBearerSentinel)) == std::wstring::npos,
               L"captured diagnostics must not contain the bearer sentinel",
               passed,
               failed);
    DebugCheck(context.captured.find(kGraphSentinel) == std::wstring::npos && context.captured.find(kUploadSentinel) == std::wstring::npos,
               L"captured diagnostics must not contain Graph or upload-session query sentinels",
               passed,
               failed);
    DebugCheck(context.captured.find(L"https/graph-api") != std::wstring::npos && context.captured.find(L"https/preauthenticated-upload") != std::wstring::npos,
               L"captured diagnostics should retain only the redacted request target classes",
               passed,
               failed);
}

struct RangedVersionSwapDebugContext final
{
    unsigned int rangeRequests     = 0u;
    bool allRangesPinned           = true;
    bool firstRangeMatchesEightMiB = true;
};

[[nodiscard]] HRESULT RangedVersionSwapDebugHook(void* cookie,
                                                 std::wstring_view method,
                                                 std::wstring_view url,
                                                 std::span<const HttpHeader> headers,
                                                 [[maybe_unused]] std::string_view bodyUtf8,
                                                 [[maybe_unused]] bool allowRetry,
                                                 HttpResponse& responseOut) noexcept
{
    auto* context = static_cast<RangedVersionSwapDebugContext*>(cookie);
    if (context == nullptr)
    {
        return E_POINTER;
    }

    if (url.find(L"graph.microsoft.com") != std::wstring_view::npos)
    {
        responseOut.statusCode = 200u;
        responseOut.body =
            R"json({"id":"swap","name":"swap.bin","size":8650752,"eTag":"\"v1\"","@microsoft.graph.downloadUrl":"https://download.test/swap","file":{}})json";
        return S_OK;
    }

    if (! OrdinalString::EqualsNoCase(method, L"GET") || url != L"https://download.test/swap")
    {
        return E_UNEXPECTED;
    }

    ++context->rangeRequests;
    context->allRangesPinned = context->allRangesPinned && FindDebugHeader(headers, L"If-Match") == L"\"v1\"";
    if (context->rangeRequests == 1u)
    {
        context->firstRangeMatchesEightMiB = FindDebugHeader(headers, L"Range") == L"bytes=0-8388607";
        responseOut.statusCode             = 206u;
        responseOut.body.assign(8u * 1024u * 1024u, 'A');
    }
    else
    {
        responseOut.statusCode = 412u;
        responseOut.body       = R"json({"error":{"code":"preconditionFailed"}})json";
    }
    return S_OK;
}

void RunDebugRangedReaderPinsVersionSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"ranged-reader version test should allocate Microsoft Drive instance", passed, failed))
    {
        return;
    }

    RangedVersionSwapDebugContext context{};
    DebugHttpRequestHookScope hook(RangedVersionSwapDebugHook, &context);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope syntheticContext(g_debugMicrosoftDriveUseSyntheticContextForSelfTest, true);

    wil::com_ptr<IFileReader> reader;
    HRESULT hr = fs->CreateFileReader(L"/@conn:microsoft-drive-selftest/swap.bin", reader.addressof());
    DebugCheck(SUCCEEDED(hr) && reader, L"ranged-reader version test should create a reader", passed, failed);
    if (FAILED(hr) || ! reader)
    {
        return;
    }

    std::vector<std::byte> buffer(8u * 1024u * 1024u);
    unsigned long bytesRead = 0u;
    hr                      = reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &bytesRead);
    DebugCheck(SUCCEEDED(hr) && bytesRead == buffer.size(), L"one 8 MiB pinned range should satisfy the bridge-sized read", passed, failed);
    DebugCheck(context.firstRangeMatchesEightMiB, L"8 MiB reader request should issue one matching range GET", passed, failed);
    buffer.resize(256u * 1024u);
    hr = reader->Read(buffer.data(), static_cast<unsigned long>(buffer.size()), &bytesRead);
    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), L"same-size version swap should fail the second range closed", passed, failed);
    DebugCheck(context.rangeRequests == 2u, L"version-swap test should issue exactly two range requests", passed, failed);
    DebugCheck(context.allRangesPinned, L"every Microsoft Drive range request should carry the metadata ETag precondition", passed, failed);
}

struct StreamingUploadDebugContext final
{
    uint64_t expectedBytes        = 0u;
    uint64_t receivedBytes        = 0u;
    unsigned int uploadPuts       = 0u;
    unsigned int cancelDeletes    = 0u;
    bool validRanges              = true;
    bool invalidNextExpectedRange = false;
};

[[nodiscard]] HRESULT StreamingUploadDebugHook(void* cookie,
                                               std::wstring_view method,
                                               std::wstring_view url,
                                               std::span<const HttpHeader> headers,
                                               std::string_view bodyUtf8,
                                               [[maybe_unused]] bool allowRetry,
                                               HttpResponse& responseOut) noexcept
{
    auto* context = static_cast<StreamingUploadDebugContext*>(cookie);
    if (context == nullptr)
    {
        return E_POINTER;
    }

    if (OrdinalString::EqualsNoCase(method, L"POST") && url.find(L"createUploadSession") != std::wstring_view::npos)
    {
        responseOut.statusCode = 200u;
        responseOut.body       = R"json({"uploadUrl":"https://sn3302.up.1drv.com/up/session?opaque=streaming-selftest"})json";
        return S_OK;
    }
    if (OrdinalString::EqualsNoCase(method, L"DELETE") && url == L"https://sn3302.up.1drv.com/up/session?opaque=streaming-selftest")
    {
        ++context->cancelDeletes;
        responseOut.statusCode = 204u;
        return S_OK;
    }
    if (! OrdinalString::EqualsNoCase(method, L"PUT") || url != L"https://sn3302.up.1drv.com/up/session?opaque=streaming-selftest")
    {
        return E_UNEXPECTED;
    }

    const uint64_t start             = context->receivedBytes;
    const uint64_t end               = start + bodyUtf8.size() - 1u;
    const std::wstring expectedRange = std::format(L"bytes {}-{}/{}", start, end, context->expectedBytes);
    context->validRanges             = context->validRanges && FindDebugHeader(headers, L"Content-Range") == expectedRange;
    context->receivedBytes += bodyUtf8.size();
    ++context->uploadPuts;
    responseOut.statusCode = context->receivedBytes == context->expectedBytes ? 201u : 202u;
    if (responseOut.statusCode == 201u)
    {
        responseOut.body = R"json({"id":"uploaded"})json";
    }
    else
    {
        const uint64_t nextOffset = context->invalidNextExpectedRange ? context->expectedBytes - 1u : context->receivedBytes;
        responseOut.body          = std::format(R"json({{"nextExpectedRanges":["{}-"]}})json", nextOffset);
    }
    return S_OK;
}

void RunDebugWriterStreamsBeforeCommitSelfTest(unsigned int& passed, unsigned int& failed)
{
    wil::com_ptr<FileSystemMicrosoftDrive> fs = MakeDebugFileSystem();
    if (! DebugCheck(static_cast<bool>(fs), L"streaming-writer test should allocate Microsoft Drive instance", passed, failed))
    {
        return;
    }

    const uint64_t chunkBytes = ComputeUploadChunkSizeBytes(fs->SnapshotSettings());
    StreamingUploadDebugContext context{.expectedBytes = chunkBytes + 100u};
    DebugHttpRequestHookScope hook(StreamingUploadDebugHook, &context);
    DebugFlagScope tokenBypass(g_debugMicrosoftDriveBypassAccessTokenForSelfTest, true);
    DebugFlagScope syntheticContext(g_debugMicrosoftDriveUseSyntheticContextForSelfTest, true);

    wil::com_ptr<IFileWriter> writer;
    HRESULT hr = fs->CreateFileWriter(L"/@conn:microsoft-drive-selftest/streamed.bin", FILESYSTEM_FLAG_ALLOW_OVERWRITE, writer.addressof());
    DebugCheck(SUCCEEDED(hr) && writer, L"streaming-writer test should create a writer", passed, failed);
    if (FAILED(hr) || ! writer)
    {
        return;
    }

    wil::com_ptr<IFileWriterExpectedSize> expectedSizeWriter;
    hr = writer->QueryInterface(IID_PPV_ARGS(expectedSizeWriter.addressof()));
    DebugCheck(SUCCEEDED(hr) && expectedSizeWriter, L"Microsoft Drive writer should expose the expected-size streaming contract", passed, failed);
    if (FAILED(hr) || ! expectedSizeWriter)
    {
        return;
    }

    hr = expectedSizeWriter->SetExpectedSize(context.expectedBytes);
    DebugCheck(SUCCEEDED(hr), L"streaming writer should accept the final size before the first Write", passed, failed);

    std::vector<std::byte> chunk(static_cast<size_t>(chunkBytes), std::byte{0x5a});
    unsigned long written = 0u;
    hr                    = writer->Write(chunk.data(), static_cast<unsigned long>(chunk.size()), &written);
    DebugCheck(SUCCEEDED(hr) && written == chunk.size(), L"streaming writer should accept and upload the first aligned chunk", passed, failed);
    DebugCheck(context.uploadPuts == 1u, L"Microsoft Drive upload must start during Write, before Commit", passed, failed);

    const std::array<std::byte, 100> tail{};
    hr = writer->Write(tail.data(), static_cast<unsigned long>(tail.size()), &written);
    DebugCheck(SUCCEEDED(hr) && written == tail.size(), L"streaming writer should upload the final short chunk", passed, failed);
    const unsigned int putsBeforeCommit = context.uploadPuts;
    hr                                  = writer->Commit();
    DebugCheck(SUCCEEDED(hr), L"streaming writer Commit should only finalize an already uploaded object", passed, failed);
    DebugCheck(context.uploadPuts == putsBeforeCommit, L"streaming writer Commit should not replay buffered file bytes", passed, failed);
    DebugCheck(context.receivedBytes == context.expectedBytes && context.validRanges,
               L"streaming upload should send every byte once with contiguous Content-Range values",
               passed,
               failed);

    writer.reset();
    DebugCheck(context.cancelDeletes == 0u, L"committed streaming upload should not cancel its completed session", passed, failed);

    context.expectedBytes            = chunkBytes + 100u;
    context.receivedBytes            = 0u;
    context.uploadPuts               = 0u;
    context.invalidNextExpectedRange = true;
    wil::com_ptr<IFileWriter> invalidRangeWriter;
    hr = fs->CreateFileWriter(L"/@conn:microsoft-drive-selftest/invalid-range.bin", FILESYSTEM_FLAG_ALLOW_OVERWRITE, invalidRangeWriter.addressof());
    wil::com_ptr<IFileWriterExpectedSize> invalidRangeExpectedSize;
    if (SUCCEEDED(hr) && invalidRangeWriter)
    {
        hr = invalidRangeWriter->QueryInterface(IID_PPV_ARGS(invalidRangeExpectedSize.addressof()));
    }
    if (SUCCEEDED(hr) && invalidRangeExpectedSize)
    {
        hr = invalidRangeExpectedSize->SetExpectedSize(context.expectedBytes);
    }
    written = 0u;
    if (SUCCEEDED(hr) && invalidRangeWriter)
    {
        hr = invalidRangeWriter->Write(chunk.data(), static_cast<unsigned long>(chunk.size()), &written);
    }
    DebugCheck(hr == HRESULT_FROM_WIN32(ERROR_INVALID_DATA), L"streaming upload should reject nextExpectedRanges beyond the submitted chunk", passed, failed);
    invalidRangeExpectedSize.reset();
    invalidRangeWriter.reset();
    context.invalidNextExpectedRange = false;
    context.receivedBytes            = 0u;
    context.uploadPuts               = 0u;
    context.cancelDeletes            = 0u;

    wil::com_ptr<IFileWriter> abandonedWriter;
    written = 0u;
    hr      = fs->CreateFileWriter(L"/@conn:microsoft-drive-selftest/abandoned.bin", FILESYSTEM_FLAG_ALLOW_OVERWRITE, abandonedWriter.addressof());
    wil::com_ptr<IFileWriterExpectedSize> abandonedExpectedSize;
    if (SUCCEEDED(hr) && abandonedWriter)
    {
        hr = abandonedWriter->QueryInterface(IID_PPV_ARGS(abandonedExpectedSize.addressof()));
    }
    if (SUCCEEDED(hr) && abandonedExpectedSize)
    {
        hr = abandonedExpectedSize->SetExpectedSize(chunkBytes + 200u);
    }
    if (SUCCEEDED(hr) && abandonedWriter)
    {
        hr = abandonedWriter->Write(tail.data(), static_cast<unsigned long>(tail.size()), &written);
    }
    DebugCheck(SUCCEEDED(hr) && abandonedWriter && abandonedExpectedSize && written == tail.size(),
               L"abandoned streaming upload should initialize and buffer a partial session",
               passed,
               failed);
    abandonedExpectedSize.reset();
    abandonedWriter.reset();
    DebugCheck(context.cancelDeletes == 1u, L"releasing an uncommitted streaming writer should cancel its upload session", passed, failed);
}
} // namespace

extern "C" __declspec(dllexport) HRESULT __stdcall RedSalamanderMicrosoftDriveDebugSelfTests(unsigned int* passed, unsigned int* failed)
{
    if (! passed || ! failed)
    {
        return E_POINTER;
    }

    *passed = 0;
    *failed = 0;

    RunDebugPatchRetryMergeSelfTest(*passed, *failed);
    RunDebugCaseOnlyRenameSelfTest(*passed, *failed);
    RunDebugDeleteRetrySelfTest(*passed, *failed);
    RunDebugPostCreateReconcileSelfTest(*passed, *failed);
    RunDebugTypeMismatchPartialMergeSelfTest(*passed, *failed);
    RunDebugDescendantMoveGuardSelfTest(*passed, *failed);
    RunDebugParseChildrenSkipsEmptyNameSelfTest(*passed, *failed);
    RunDebugEmptyNameChildBlocksSourceDeleteSelfTest(*passed, *failed);
    RunDebugMergeNeverRecursivelyDeletesSourceFolderSelfTest(*passed, *failed);
    RunDebugMoveRejectsInvalidOptionsSizeSelfTest(*passed, *failed);
    RunDebugMergeRecursionDepthCapSelfTest(*passed, *failed);
    RunDebugOverwriteCancelRestoresDestinationSelfTest(*passed, *failed);
    RunDebugCommittedMoveCleanupFailureIsWarningSelfTest(*passed, *failed);
    RunDebugCredentialUrlValidationSelfTest(*passed, *failed);
    RunDebugContinuationBoundarySelfTest(*passed, *failed);
    RunDebugCredentialDiagnosticRedactionSelfTest(*passed, *failed);
    RunDebugRangedReaderPinsVersionSelfTest(*passed, *failed);
    RunDebugWriterStreamsBeforeCommitSelfTest(*passed, *failed);

    return *failed == 0u ? S_OK : E_FAIL;
}
#endif
